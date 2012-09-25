Автоматические тесты DSK
---
Automatic SDK tests
==============

Для настройки тестового стенда необходимо:
 - В настройках безопасности Internet explorer разрешить выполнение активных скриптов и запуск элементов ActiveX, не помеченных как безопасные
 - В хранилище личных сертификатов должен быть хотябы один сертификат
 - В КриптоАРМ-е должна быть задана настройка по умолчанию, в которой должны быть определены работоспособные параметры подписывания и шифрования (ключ сертификата не должен быть защищен пин кодом)

Так же:
 - При смене версии КриптоАРМ-а необходимо внести исправления в коде проверки функции GetDigtCryptoVersion - внести версию без номера билда
 - Если первый возвращаемый из хранилища сертификат должен иметь статус "корректен". Иначе необходимо внести ожидаемый статус в коде проверки функции VerifyCertificateStatus (первый возвращаемый в свойстве Store объектом CertificateStore)

---

For setting up test stand:
 - In the Internet Explorer security level settings enable options "Initialize and script ActiveX controls not marked as safe" and "Active scripting"
 - Must be at least one certificate in personal certificate store
 - In the CryptoARM must be defined default operational profile with operable signing and encrypting settings (private key of the certificate must not be protected by PIN-code)


Also:
 - For a new CryptoARM version necessary to update expected version in the test code of the "GetDigtCryptoVersion" function
 - If first certificate in the private certificates store have not a "correct" status, necessary to correct expected status in the test code of the "VerifyCertificateStatus" function (first certificate that returned with a property "Store" of an object "CertificateStore")

