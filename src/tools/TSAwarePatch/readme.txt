TSAwarePatch by Alex Dragokas - Terminal server awareness patch 

Used to add PE flags for supporting ASLR, DEP and avoiding differrent kinds of redirections on server systems.
Also, it patch MajorSubsystemVersion and MajorOperatingSystemVersion fields to revert 4.0 subsystem in order to support Win 2k/XP just in case you used new version of linker from modern Visual Studio.

Usage:

TSAwarePatch.exe [path to exe]

----------

HiJackThis note:

This program is not included in HiJackThis Fork resources.
However, it is used for building its binary.

------------
Checksum:

TSAwarePatch.exe
Digitally signed by Stanislav Polshyn.

Certificate's thumbprint should be: 1b78ef517e81a07d1c1c4c6adfa66a2b7c3269c3
Serial number is: 31f8f5fb790c592476ce0f3320dc4af1
