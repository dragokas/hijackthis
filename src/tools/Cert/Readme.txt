List of Untrusted and Trusted Root certificates

This folder contains last version of certificates included in HiJackThis in form of fingerprint (hashes).

Using:

Use _11_Check_New_Certificates.cmd script (from the root source folder) to compare and check for new certificates.

If Microsoft will release:

1) New own root certificate, you need to add its fingerprint to modVerifyDigiSign.bas::IsMicrosoftCertHash function.

2) New disallowed certificate, you need to add its fingerprint to modMain.bas::LoadStuff => colSafeCert() collection.

------------

* HiJackThis note:

These files are not included in HiJackThis Fork resources.
However, it is used to track new certificates.

It is good idea to run this script at least once a month.