"""Sanitize common connection-string, authorization-header, and signed-URL secrets."""
import re


_SECRET_FIELD = (
    r'[A-Za-z0-9_.-]*(?:password|pwd|token|secret|credential|signature|'
    r'api[_ -]?key|private[_ -]?key|access[_ -]?key|signing[_ -]?key|'
    r'encryption[_ -]?key|subscription[_ -]?key|code[_ -]?verifier|'
    r'shared[_ -]?access[_ -]?(?:signature|key))[A-Za-z0-9_.-]*'
)
_ASSIGNMENT = re.compile(
    r'(?i)(?<![A-Za-z0-9_.-])(["\']?' + _SECRET_FIELD + r'["\']?\s*[:=]\s*)'
    r'("[^"\r\n]*"|\'[^\'\r\n]*\'|[^;,&\s]+)'
)
_URI_PASSWORD = re.compile(r'(?i)(://[^:/@\s]+:)[^@/\s]+(@)')
_AUTHORIZATION = re.compile(
    r'(?im)(\b(?:proxy-)?authorization\b["\']?\s*[:=]\s*)'
    r'("[^"\r\n]*"|\'[^\'\r\n]*\'|[^\r\n]+)'
)
_SIGNED_QUERY_VALUE = re.compile(
    r'(?i)([?&](?:access[_-]?token|token|api[_-]?key|client[_-]?secret|'
    r'subscription[_-]?key|sig|signature|x-amz-(?:signature|credential|security-token)|'
    r'x-goog-(?:signature|credential)|password|pwd|secret|credential)=)'
    r'([^&#\s]+)'
)


def safe_error(value):
    text = str(value)
    text = _ASSIGNMENT.sub(r'\1[redacted]', text)
    text = _AUTHORIZATION.sub(r'\1[redacted]', text)
    text = _SIGNED_QUERY_VALUE.sub(r'\1[redacted]', text)
    return _URI_PASSWORD.sub(r'\1[redacted]\2', text)
