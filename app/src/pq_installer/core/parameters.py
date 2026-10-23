"""Conservative editing for literal Power Query parameter expressions."""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
import re


@dataclass(frozen=True)
class ParameterLiteral:
    kind: str
    value: str | bool | Decimal
    start: int
    end: int


def _skip_trivia(text, position):
    while position < len(text):
        if text[position].isspace():
            position += 1
        elif text.startswith('//', position):
            newline = text.find('\n', position + 2)
            position = len(text) if newline < 0 else newline + 1
        elif text.startswith('/*', position):
            close = text.find('*/', position + 2)
            if close < 0:
                return len(text) + 1
            position = close + 2
        else:
            break
    return position


def _parse_text(text, position):
    start = position
    position += 1
    value = []
    while position < len(text):
        if text[position] != '"':
            value.append(text[position]); position += 1
        elif position + 1 < len(text) and text[position + 1] == '"':
            value.append('"'); position += 2
        else:
            return ''.join(value), start, position + 1
    return None


def _parse_literal(text, position):
    if position >= len(text):
        return None
    if text[position] == '"':
        parsed = _parse_text(text, position)
        return ('Text', *parsed) if parsed else None
    number = re.match(r'[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?', text[position:])
    if number:
        token = number.group(0)
        try:
            value = Decimal(token)
        except InvalidOperation:
            return None
        if not value.is_finite():
            return None
        return 'Number', value, position, position + len(token)
    logical = re.match(r'(?i:true|false)(?![\w])', text[position:])
    if logical:
        token = logical.group(0)
        return 'Logical', token.casefold() == 'true', position, position + len(token)
    return None


def _balanced_metadata_end(text, start):
    depth = 0
    position = start
    while position < len(text):
        if text.startswith('//', position):
            newline = text.find('\n', position + 2)
            position = len(text) if newline < 0 else newline + 1
            continue
        if text.startswith('/*', position):
            close = text.find('*/', position + 2)
            if close < 0:
                return None
            position = close + 2
            continue
        if text[position] == '"':
            quoted = _parse_text(text, position)
            if not quoted:
                return None
            position = quoted[2]
            continue
        if text[position] == '[':
            depth += 1
        elif text[position] == ']':
            depth -= 1
            if depth == 0:
                return position
            if depth < 0:
                return None
        position += 1
    return None


def _has_parameter_metadata(metadata):
    props = {}
    position = 0
    while position < len(metadata):
        position = _skip_trivia(metadata, position)
        if position >= len(metadata):
            break
        name = re.match(r'[A-Za-z_][\w]*', metadata[position:])
        if not name:
            return None
        key = name.group(0).casefold()
        position += len(name.group(0))
        position = _skip_trivia(metadata, position)
        if position >= len(metadata) or metadata[position] != '=':
            return None
        position = _skip_trivia(metadata, position + 1)
        literal = _parse_literal(metadata, position)
        if not literal:
            return None
        value_kind, value, _, position = literal
        props[key] = value
        position = _skip_trivia(metadata, position)
        if position == len(metadata):
            break
        if metadata[position] != ',':
            return None
        position += 1
    if props.get('isparameterquery') is not True:
        return None
    kind = props.get('type')
    return kind if isinstance(kind, str) else None


def inspect_parameter(m_code):
    """Return the supported declared literal, or None when syntax is ambiguous."""
    if not isinstance(m_code, str):
        return None
    start = _skip_trivia(m_code, 0)
    literal = _parse_literal(m_code, start)
    if not literal:
        return None
    kind, value, value_start, value_end = literal
    position = _skip_trivia(m_code, value_end)
    match = re.match(r'(?i:meta)(?![\w])', m_code[position:])
    if not match:
        return None
    position = _skip_trivia(m_code, position + len(match.group(0)))
    if position >= len(m_code) or m_code[position] != '[':
        return None
    close = _balanced_metadata_end(m_code, position)
    if close is None:
        return None
    declared_type = _has_parameter_metadata(m_code[position + 1:close])
    if not declared_type or declared_type.casefold() != kind.casefold():
        return None
    if _skip_trivia(m_code, close + 1) != len(m_code):
        return None
    return ParameterLiteral(kind, value, value_start, value_end)


def _encode_value(kind, value):
    if kind == 'Text':
        if not isinstance(value, str):
            raise ValueError('Text parameters require a text value.')
        return '"' + value.replace('"', '""') + '"'
    if kind == 'Logical':
        if isinstance(value, bool):
            return 'true' if value else 'false'
        if isinstance(value, str) and value.strip().casefold() in {'true', 'false'}:
            return value.strip().casefold()
        raise ValueError('Logical parameters accept only true or false.')
    if kind == 'Number':
        if isinstance(value, bool):
            raise ValueError('Number parameters require a finite number.')
        try:
            number = Decimal(str(value).strip())
        except (InvalidOperation, ValueError):
            raise ValueError('Number parameters require a finite number.') from None
        if not number.is_finite() or abs(number.adjusted()) > 308:
            raise ValueError('Number parameters require a finite number within the supported range.')
        encoded = format(number.normalize(), 'f')
        if len(encoded) > 256:
            raise ValueError('Number parameter is too long.')
        return encoded
    raise ValueError(f'Parameter type {kind!r} is read-only.')


def edit_parameter(m_code, value):
    """Replace only a recognized parameter literal; preserve every other byte of text."""
    parameter = inspect_parameter(m_code)
    if parameter is None:
        raise ValueError('This expression is not a supported declared literal parameter; it remains read-only.')
    literal = _encode_value(parameter.kind, value)
    return m_code[:parameter.start] + literal + m_code[parameter.end:]


def edit_parameter_file(source_path, value, *, destination):
    """Write an edited copy, preserving UTF-8 BOM and all non-literal source bytes."""
    source_path = Path(source_path)
    destination = Path(destination)
    if destination.exists():
        raise FileExistsError(f'Refusing to overwrite an existing source file: {destination}')
    raw = source_path.read_bytes()
    bom = raw.startswith(b'\xef\xbb\xbf')
    payload = raw[3:] if bom else raw
    try:
        source = payload.decode('utf-8')
    except UnicodeDecodeError as exc:
        raise ValueError('Parameter source must be valid UTF-8.') from exc
    edited = edit_parameter(source, value)
    destination.parent.mkdir(parents=True, exist_ok=True)
    data = (b'\xef\xbb\xbf' if bom else b'') + edited.encode('utf-8')
    with destination.open('xb') as stream:
        stream.write(data)
    return destination
