"""Token-aware mapping for literal arguments of common M source functions."""

from __future__ import annotations

from dataclasses import dataclass
import re


_SOURCE_FUNCTIONS = frozenset({
    ('file', 'contents'), ('folder', 'files'), ('folder', 'contents'),
    ('sharepoint', 'files'), ('sharepoint', 'contents'), ('web', 'contents'),
})


@dataclass(frozen=True)
class MSourceChange:
    start: int
    end: int
    source_function: str
    original: str
    replacement: str
    mapping_key: str


@dataclass(frozen=True)
class MSourceMappingPreview:
    original: str
    mapped: str
    changes: tuple[MSourceChange, ...]
    warnings: tuple[str, ...]


@dataclass(frozen=True)
class _Token:
    kind: str
    value: str
    start: int
    end: int


def _tokens(text):
    result = []
    position = 0
    while position < len(text):
        if text[position].isspace():
            position += 1
            continue
        if text.startswith('//', position):
            newline = text.find('\n', position + 2)
            position = len(text) if newline < 0 else newline + 1
            continue
        if text.startswith('/*', position):
            depth = 1
            cursor = position + 2
            while cursor < len(text) and depth:
                if text.startswith('/*', cursor):
                    depth += 1; cursor += 2
                elif text.startswith('*/', cursor):
                    depth -= 1; cursor += 2
                else:
                    cursor += 1
            position = cursor
            continue
        start = position
        if text[position] == '"':
            position += 1
            value = []
            while position < len(text):
                if text[position] != '"':
                    value.append(text[position]); position += 1
                elif position + 1 < len(text) and text[position + 1] == '"':
                    value.append('"'); position += 2
                else:
                    position += 1
                    result.append(_Token('string', ''.join(value), start, position))
                    break
            else:
                result.append(_Token('invalid_string', '', start, len(text)))
            continue
        if text.startswith('#"', position):
            position += 2
            while position < len(text):
                if text[position] == '"':
                    if position + 1 < len(text) and text[position + 1] == '"':
                        position += 2
                    else:
                        position += 1; break
                else:
                    position += 1
            result.append(_Token('quoted_identifier', '', start, position))
            continue
        identifier = re.match(r'[A-Za-z_][A-Za-z0-9_]*', text[position:])
        if identifier:
            value = identifier.group(0)
            position += len(value)
            result.append(_Token('identifier', value, start, position))
            continue
        result.append(_Token('punctuation', text[position], start, start + 1))
        position += 1
    return result


def _normalize_path(value):
    normalized = value.replace('\\', '/')
    if len(normalized) > 1 and normalized.endswith('/'):
        normalized = normalized.rstrip('/')
    return normalized


def _mapped_path(value, mappings):
    matches = []
    normalized_value = _normalize_path(value)
    for mapping in mappings:
        key, source_root, target_root = mapping['key'], mapping['from'], mapping['to']
        source_normalized = _normalize_path(source_root)
        candidate_normalized = _normalize_path(target_root)
        if not source_normalized:
            continue
        if not normalized_value.casefold().startswith(source_normalized.casefold()):
            continue
        boundary = len(source_normalized)
        if (boundary < len(normalized_value) and not source_normalized.endswith('/')
                and normalized_value[boundary] != '/'):
            continue
        suffix = value[len(source_root.rstrip('\\/')):]
        if source_root.endswith(('\\', '/')) and suffix and suffix[0] not in ('\\', '/'):
            suffix = ('\\' if '\\' in source_root else '/') + suffix
        replacement = target_root.rstrip('\\/') + suffix
        matches.append((len(source_normalized), key, replacement, candidate_normalized))
    if not matches:
        return None, None
    specificity = max(row[0] for row in matches)
    best = [row for row in matches if row[0] == specificity]
    targets = {_normalize_path(row[2]).casefold() for row in best}
    if len(targets) > 1:
        return None, 'ambiguous'
    _, key, replacement, _ = sorted(best, key=lambda row: row[1].casefold())[0]
    return replacement, key


def _encode_m_text(value):
    return '"' + value.replace('"', '""') + '"'


def preview_m_source_mapping(m_code, mappings):
    """Map only literal first arguments of recognized source functions.

    Comments, strings outside a source call, identifiers, and dynamic expressions
    are never rewritten. Dynamic source expressions produce a warning for review.
    """
    if not isinstance(m_code, str):
        raise TypeError('M source must be text.')
    active = []
    for mapping in mappings:
        if mapping.get('enabled', True):
            key, source, target = mapping.get('key'), mapping.get('from'), mapping.get('to')
            if not all(isinstance(value, str) and value for value in (key, source, target)):
                raise ValueError('M source mapping requires non-empty key, from, and to values.')
            active.append({'key': key, 'from': source, 'to': target})
    if not active:
        return MSourceMappingPreview(m_code, m_code, (), ())
    tokens = _tokens(m_code)
    changes = []
    warnings = []
    for index in range(len(tokens) - 3):
        first, dot, second, opening = tokens[index:index + 4]
        if (first.kind != 'identifier' or dot.value != '.' or second.kind != 'identifier'
                or opening.value != '(' or (first.value.casefold(), second.value.casefold()) not in _SOURCE_FUNCTIONS):
            continue
        function = f'{first.value}.{second.value}'
        argument_index = index + 4
        if argument_index >= len(tokens) or tokens[argument_index].value == ')':
            warnings.append(f'{function} at character {first.start}: missing or dynamic source argument left unchanged.')
            continue
        argument = tokens[argument_index]
        if argument.kind != 'string':
            warnings.append(f'{function} at character {first.start}: dynamic or non-literal source expression left unchanged.')
            continue
        replacement, mapping_key = _mapped_path(argument.value, active)
        if mapping_key == 'ambiguous':
            warnings.append(f'{function} at character {first.start}: source mapping is ambiguous; literal left unchanged.')
        elif replacement is not None and replacement != argument.value:
            changes.append(MSourceChange(argument.start, argument.end, function,
                                          argument.value, replacement, mapping_key))
    mapped = m_code
    for change in sorted(changes, key=lambda item: item.start, reverse=True):
        mapped = mapped[:change.start] + _encode_m_text(change.replacement) + mapped[change.end:]
    return MSourceMappingPreview(m_code, mapped, tuple(sorted(changes, key=lambda item: item.start)), tuple(warnings))
