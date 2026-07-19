# CivilEnginStructuresOpenClaw

Это независимая копия проекта `CivilEnginStructures`.

Внутренние имена сборки и пространства имен намеренно оставлены прежними:

- сборка: `CivilEnginStructures.dll`;
- класс загрузчика: `CivilEnginStructures.RopExample1.RopExample1PluginHost`;
- манифест: `InfrastradaBridgeTools.plugin`.

Их переименование нарушило бы связь между DLL и штатным манифестом Topomatic Robur.

## Сборка

Конфигурация: `Debug`, .NET Framework 4.8.

Результат сборки находится в `bin\Debug`.

## Установка

Пакет `CivilEnginStructuresOpenClaw.tpm` устанавливается штатной программой
`C:\Program Files\Topomatic Robur Road 16.0\TopomaticPackageManager.exe`.

Установка требует прав администратора, поскольку файлы Robur находятся в
`C:\Program Files`.

Перед первой попыткой подключения исходные установленные файлы плагина были
скопированы в `OpenClawDeploymentBackup\2026-07-19-before-first-deploy`.
