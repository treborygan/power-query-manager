"""Bounded M reference scanner. It is intentionally not a full M parser."""
from dataclasses import dataclass
import unicodedata


_M_IDENTIFIER_START_CATEGORIES={'Lu','Ll','Lt','Lm','Lo','Nl'}
_M_IDENTIFIER_PART_CATEGORIES=_M_IDENTIFIER_START_CATEGORIES|{'Nd','Pc','Mn','Mc','Cf'}
_M_CONTROL_ESCAPES={'cr':'\r','lf':'\n','tab':'\t'}
_M_HASH_KEYWORDS={'binary','date','datetime','datetimezone','duration','infinity','nan','sections','shared','table','time'}


def _is_m_identifier_start(character):
    return character=='_' or unicodedata.category(character) in _M_IDENTIFIER_START_CATEGORIES


def _is_m_identifier_part(character):
    return unicodedata.category(character) in _M_IDENTIFIER_PART_CATEGORIES


def _decode_m_character_escape(text, start):
    """Decode one M #(...) escape, returning (value, next_index, valid)."""
    if not text.startswith('#(',start):return '',start,False
    end=text.find(')',start+2)
    if end<0:return '',len(text),False
    decoded=[]
    for escape in text[start+2:end].split(','):
        if escape in _M_CONTROL_ESCAPES:
            decoded.append(_M_CONTROL_ESCAPES[escape]);continue
        if escape=='#':
            decoded.append('#');continue
        if len(escape) not in (4,8) or any(character not in '0123456789abcdefABCDEF' for character in escape):
            return '',end+1,False
        codepoint=int(escape,16)
        if codepoint>0x10FFFF or 0xD800<=codepoint<=0xDFFF:
            return '',end+1,False
        decoded.append(chr(codepoint))
    return ''.join(decoded),end+1,True


@dataclass(frozen=True)
class ReferenceAnalysis:
    dependencies: tuple[str, ...]
    ambiguous: bool
    reason: str = ''


def _tokens(text):
    tokens=[];i=0;ambiguous=False
    while i<len(text):
        c=text[i]
        if c.isspace():i+=1;continue
        if text.startswith('//',i):
            end=i+2
            while end<len(text) and text[end] not in '\r\n\u0085\u2028\u2029':end+=1
            if end>=len(text):i=end
            elif text.startswith('\r\n',end):i=end+2
            else:i=end+1
            continue
        if text.startswith('/*',i):
            end=text.find('*/',i+2)
            if end<0:return tokens,True
            i=end+2;continue
        if c=='"':
            i+=1
            while i<len(text):
                if text[i]=='"':
                    if i+1<len(text) and text[i+1]=='"':i+=2;continue
                    i+=1;break
                i+=1
            else:ambiguous=True
            tokens.append(('string',''));continue
        if text.startswith('#"',i):
            i+=2;value=[];closed=False;valid=True
            while i<len(text):
                if text[i]=='"':
                    if i+1<len(text) and text[i+1]=='"':value.append('"');i+=2;continue
                    i+=1;closed=True;break
                if text.startswith('#(',i):
                    decoded,next_index,escape_valid=_decode_m_character_escape(text,i)
                    if escape_valid:value.append(decoded)
                    else:valid=False
                    i=next_index;continue
                value.append(text[i]);i+=1
            if not closed:ambiguous=True
            if not valid:ambiguous=True
            tokens.append(('quoted_identifier',''.join(value) if valid else None));continue
        if c=='#':
            if i+1<len(text) and _is_m_identifier_start(text[i+1]):
                start=i+1;i=start+1
                while i<len(text) and _is_m_identifier_part(text[i]):i+=1
                value=text[start:i]
                if value not in _M_HASH_KEYWORDS:ambiguous=True
                tokens.append(('hash_keyword',value));continue
            ambiguous=True
            tokens.append(('punct','#'));i+=1;continue
        if _is_m_identifier_start(c):
            start=i;i+=1
            while i<len(text) and _is_m_identifier_part(text[i]):i+=1
            while (i+1<len(text) and text[i]=='.' and
                   _is_m_identifier_start(text[i+1])):
                i+=2
                while i<len(text) and _is_m_identifier_part(text[i]):i+=1
            value=text[start:i]
            kind='qualified_identifier' if '.' in value else 'identifier'
            tokens.append((kind,value));continue
        tokens.append(('punct',c));i+=1
    return tokens,ambiguous


def analyze_query_references(formula, query_names):
    tokens,lexical_ambiguity=_tokens(str(formula))
    # M identifiers are case-sensitive. Folding here can create false edges
    # and can make a differently-cased local binding hide a real dependency.
    names={str(name):str(name) for name in query_names}
    square_depth=[];depth=0;delimiters=[];malformed=False
    for token in tokens:
        if token==('punct','['):depth+=1
        square_depth.append(depth)
        if token==('punct',']'):depth=max(0,depth-1)
        if token[0]=='punct' and token[1] in '([{':delimiters.append(token[1])
        elif token[0]=='punct' and token[1] in ')]}':
            expected={')':'(',']':'[','}':'{'}[token[1]]
            if not delimiters or delimiters.pop()!=expected:malformed=True
    malformed=malformed or bool(delimiters)
    dynamic=False
    unresolved_quoted=[]
    for i,(kind,value) in enumerate(tokens):
        if kind=='qualified_identifier' and value in {'Expression.Evaluate','Record.Field'}:dynamic=True
        if kind=='hash_keyword' and value in {'sections','shared'}:dynamic=True
        if kind=='identifier' and value=='Evaluate' and i>=2 and tokens[i-1]==('punct','.') and tokens[i-2]==('identifier','Expression'):dynamic=True
        if kind=='identifier' and value=='Field' and i>=2 and tokens[i-1]==('punct','.') and tokens[i-2]==('identifier','Record'):dynamic=True
    deps=[];shadowed=set();ambiguous=lexical_ambiguity or dynamic or malformed
    for i,(kind,value) in enumerate(tokens):
        if kind not in {'identifier','qualified_identifier','quoted_identifier'}:continue
        key=value
        if kind=='quoted_identifier' and key is not None and key not in names and key not in unresolved_quoted:
            unresolved_quoted.append(key)
        if key not in names:continue
        previous=tokens[i-1] if i else None
        following=tokens[i+1] if i+1<len(tokens) else None
        # Binding names and record field labels are declarations, not query references.
        if following==('punct','='):
            if square_depth[i]>0:continue
            shadowed.add(key);continue
        if previous==('punct','[') and following==('punct',']'):
            continue
        if previous==('punct','.'):
            continue
        if following==('punct','=>'):
            shadowed.add(key);continue
        if key not in shadowed and names[key] not in deps:deps.append(names[key])
    deps=[name for name in deps if name not in shadowed]
    if unresolved_quoted:ambiguous=True
    for key in shadowed:
        if key in names:ambiguous=True
    reasons=[]
    if lexical_ambiguity:reasons.append('invalid or incomplete lexical token or comment')
    if dynamic:reasons.append('dynamic or #shared reference construct')
    if malformed:reasons.append('unbalanced M delimiters')
    if shadowed:reasons.append('local binding shadows a query name')
    if unresolved_quoted:reasons.append('unresolved quoted identifier: '+', '.join(unresolved_quoted))
    return ReferenceAnalysis(tuple(deps),ambiguous,'; '.join(reasons))


def find_dependency_cycles(graph):
    """Return directed cycle paths; absent/ambiguous edges should be omitted by caller."""
    nodes=set(graph);state={};stack=[];found=[]
    def visit(node):
        state[node]=1;stack.append(node)
        for neighbor in sorted(set(graph.get(node,()))):
            if neighbor not in nodes:continue
            if state.get(neighbor)==1:
                cycle=stack[stack.index(neighbor):]+[neighbor]
                body=cycle[:-1]
                rotations=[body[i:]+body[:i] for i in range(len(body))]
                normalized=min(tuple(x) for x in rotations)
                found.append(normalized+(normalized[0],))
            elif state.get(neighbor,0)==0:visit(neighbor)
        stack.pop();state[node]=2
    for node in sorted(nodes):
        if state.get(node,0)==0:visit(node)
    return tuple(sorted(set(found)))


def analyze_import_dependencies(sources):
    """Analyze source rows for advisory direct references; never changes their order."""
    from .query_loader import read_m_code
    names=[source.name for source in sources]
    graph={source.name:[] for source in sources if source.enabled}
    analyses={};disabled={}
    for source in sources:
        if not source.enabled:
            disabled[source.name]=source.name;continue
        try:analysis=analyze_query_references(read_m_code(source),names)
        except Exception as exc:
            analyses[source.name]=(None,str(exc));continue
        analyses[source.name]=(analysis,'')
        if not analysis.ambiguous:
            graph[source.name]=[name for name in analysis.dependencies if name in graph]
    cycles=find_dependency_cycles(graph)
    cycle_for={name:cycle for cycle in cycles for name in cycle[:-1]}
    result={}
    for source in sources:
        if not source.enabled:
            result[source.name]='Disabled query; dependent imports should review this source.';continue
        analysis,error=analyses.get(source.name,(None,'unknown'))
        if analysis is None:
            result[source.name]=f'Dependency analysis unavailable: {error}';continue
        disabled_refs=[name for name in analysis.dependencies if name in disabled]
        if analysis.ambiguous:
            result[source.name]='Ambiguous M reference' + (f': {analysis.reason}' if analysis.reason else '')
        elif disabled_refs:
            result[source.name]='Disabled dependency: '+', '.join(disabled_refs)
        elif source.name in cycle_for:
            result[source.name]='Cycle: '+' -> '.join(cycle_for[source.name])
        elif analysis.dependencies:
            result[source.name]='Depends on: '+', '.join(analysis.dependencies)
        else:
            result[source.name]='No direct query references detected'
    return result
