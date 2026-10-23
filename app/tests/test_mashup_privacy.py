import hashlib
import struct
import zipfile

import pytest

from pq_installer.core.mashup_privacy import WorkbookPrivacy, read_workbook_privacy


def _mashup(parts, permissions, bindings=b'ciphertext'):
    fields = (parts, permissions, b'<Metadata/>', bindings)
    return struct.pack('<I', 0) + b''.join(
        struct.pack('<I', len(field)) + field for field in fields
    )


def _permission_xml(firewall='true', group=None):
    group_xml = '<WorkbookGroupType xsi:nil="true" />' if group is None else f'<WorkbookGroupType>{group}</WorkbookGroupType>'
    return (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<PermissionList xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<CanEvaluateFuturePackages>false</CanEvaluateFuturePackages>'
        f'<FirewallEnabled>{firewall}</FirewallEnabled>{group_xml}</PermissionList>'
    ).encode('utf-8')


def _bound_plaintext(parts, permissions):
    return (struct.pack('<I', 32) + hashlib.sha256(parts).digest()
            + struct.pack('<I', 32) + hashlib.sha256(permissions).digest())


def _xlsx(path, mashup=None):
    with zipfile.ZipFile(path, 'w') as archive:
        if mashup is not None:
            import base64
            archive.writestr('customXml/item1.xml',
                '<DataMashup xmlns="http://schemas.microsoft.com/DataMashup">'
                + base64.b64encode(mashup).decode('ascii') + '</DataMashup>')
        else:
            archive.writestr('[Content_Types].xml', '<Types/>')


def test_read_workbook_privacy_extracts_bound_firewall_and_group_type(tmp_path):
    parts = b'package-parts'
    permissions = _permission_xml('true', 'SeparatePrivate')
    binary = _mashup(parts, permissions)
    workbook = tmp_path / 'private.xlsx'
    _xlsx(workbook, binary)

    result = read_workbook_privacy(
        workbook,
        dpapi_unprotect=lambda _ciphertext, entropy: (
            _bound_plaintext(parts, permissions)
            if entropy == b'DataExplorer Package Components' else b''
        ),
    )

    assert result == WorkbookPrivacy(firewall_enabled=True, workbook_group_type='SeparatePrivate')


def test_read_workbook_privacy_rejects_permissions_with_stale_binding(tmp_path):
    parts = b'package-parts'
    permissions = _permission_xml('true')
    binary = _mashup(parts, permissions)
    workbook = tmp_path / 'tampered.xlsx'
    _xlsx(workbook, binary)

    with pytest.raises(ValueError, match='binding'):
        read_workbook_privacy(workbook, dpapi_unprotect=lambda *_: b'invalid')


def test_read_workbook_privacy_returns_none_when_workbook_has_no_mashup(tmp_path):
    workbook = tmp_path / 'empty.xlsx'
    _xlsx(workbook)

    assert read_workbook_privacy(workbook) is None


def test_read_workbook_privacy_accepts_specified_unbound_permissions(tmp_path):
    permissions = _permission_xml('true', 'Organizational')
    workbook = tmp_path / 'unbound.xlsx'
    _xlsx(workbook, _mashup(b'package-parts', permissions, bindings=b'\x00'))

    assert read_workbook_privacy(workbook) == WorkbookPrivacy(True, 'Organizational')


def test_read_workbook_privacy_rejects_unknown_permission_values(tmp_path):
    parts = b'package-parts'
    permissions = _permission_xml('false', 'Unknown')
    binary = _mashup(parts, permissions)
    workbook = tmp_path / 'unknown.xlsx'
    _xlsx(workbook, binary)

    with pytest.raises(ValueError, match='WorkbookGroupType'):
        read_workbook_privacy(
            workbook,
            dpapi_unprotect=lambda _ciphertext, _entropy: _bound_plaintext(parts, permissions),
        )


def test_read_workbook_privacy_rejects_duplicate_fields(tmp_path):
    parts = b'package-parts'
    permissions = (
        '<PermissionList><FirewallEnabled>true</FirewallEnabled>'
        '<FirewallEnabled>false</FirewallEnabled></PermissionList>'
    ).encode('utf-8')
    workbook = tmp_path / 'duplicate.xlsx'
    _xlsx(workbook, _mashup(parts, permissions))

    with pytest.raises(ValueError, match='duplicate fields'):
        read_workbook_privacy(
            workbook,
            dpapi_unprotect=lambda _ciphertext, _entropy: _bound_plaintext(parts, permissions),
        )


def test_read_workbook_privacy_rejects_document_type_declarations(tmp_path):
    workbook = tmp_path / 'doctype.xlsx'
    with zipfile.ZipFile(workbook, 'w') as archive:
        archive.writestr(
            'customXml/item1.xml',
            '<!DOCTYPE DataMashup [<!ENTITY privacy "true">]><DataMashup>&privacy;</DataMashup>',
        )

    with pytest.raises(ValueError, match='prohibited document declarations'):
        read_workbook_privacy(workbook)
