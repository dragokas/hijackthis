English:

File alignment by 4-bytes boundary.

Used to patch file before adding to resources of project.
<NUL> characters (ASCII = 0) will be added to the end of file.

As known, unlike IDE mode, compiled program automatically adds bytes by 4-byte boundary
during loading resource by LoadResData function.
To avoid these random data in the end of resource, we append <NUL> on our own.

-------------------------------------------------------------------------------

Russian:

Выравнивание файла по 4-байтовой границе.

Используется для модификации файла перед внесением его в ресурсы проекта.
В конец файла дописываются знаки <NUL> (ASCII = 0).

Как известно при загрузке ресурса через LoadResData скомпилированное приложение
в отличие от режима IDE автоматически дописывает байты до 4-байтовой границы.
Чтобы избечь случайных данных в конце ресурса, дописываем <NUL> самостоятельно.
