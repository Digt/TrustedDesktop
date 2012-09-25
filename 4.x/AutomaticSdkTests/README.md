�������������� ����� DSK
---
Automatic SDK tests
==============

��� ��������� ��������� ������ ����������:
 - � ���������� ������������ Internet explorer ��������� ���������� �������� �������� � ������ ��������� ActiveX, �� ���������� ��� ����������
 - � ��������� ������ ������������ ������ ���� ������ ���� ����������
 - � ���������-� ������ ���� ������ ��������� �� ���������, � ������� ������ ���� ���������� ��������������� ��������� ������������ � ���������� (���� ����������� �� ������ ���� ������� ��� �����)

��� ��:
 - ��� ����� ������ ���������-� ���������� ������ ����������� � ���� �������� ������� GetDigtCryptoVersion - ������ ������ ��� ������ �����
 - ���� ������ ������������ �� ��������� ���������� ������ ����� ������ "���������". ����� ���������� ������ ��������� ������ � ���� �������� ������� VerifyCertificateStatus (������ ������������ � �������� Store �������� CertificateStore)

---

For setting up test stand:
 - In the Internet Explorer security level settings enable options "Initialize and script ActiveX controls not marked as safe" and "Active scripting"
 - Must be at least one certificate in personal certificate store
 - In the CryptoARM must be defined default operational profile with operable signing and encrypting settings (private key of the certificate must not be protected by PIN-code)


Also:
 - For a new CryptoARM version necessary to update expected version in the test code of the "GetDigtCryptoVersion" function
 - If first certificate in the private certificates store have not a "correct" status, necessary to correct expected status in the test code of the "VerifyCertificateStatus" function (first certificate that returned with a property "Store" of an object "CertificateStore")

