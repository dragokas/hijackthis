English:

File alignment by 4-bytes boundary.

Used to patch file before adding to resources of project.
<NUL> characters (ASCII = 0) will be added to the end of file.

As known, unlike IDE mode, compiled program automatically adds bytes by 4-byte boundary
during loading resource by LoadResData function.
To avoid these random data in the end of resource, we append <NUL> on our own.

Provided as backward compatibility with VB6 pre-SP6 IDE update.

-------------------------------------------------------------------------------

Russian:

Выравнивание файла по 4-байтовой границе.

Используется для модификации файла перед внесением его в ресурсы проекта.
В конец файла дописываются знаки <NUL> (ASCII = 0).

Как известно при загрузке ресурса через LoadResData скомпилированное приложение
в отличие от режима IDE автоматически дописывает байты до 4-байтовой границы.
Чтобы избежать случайных данных в конце ресурса, дописываем <NUL> самостоятельно.

Предоставлен в качестве обратной совместимости с VB6 IDE без обновления SP6.

------------

* HiJackThis note:

This program is not included in HiJackThis Fork resources.
However, it is used to finalize project building.

------------
Checksum:

Align4byte.exe

Digitally signed by Stanislav Polshyn.

Certificate's thumbprint should be: 1b78ef517e81a07d1c1c4c6adfa66a2b7c3269c3
Serial number is: 31f8f5fb790c592476ce0f3320dc4af1
