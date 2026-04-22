# Windows Auto-Deploy for Exam Computers

Автоматическая установка и настройка Windows 11 для учебных компьютеров (кабинет информатики, олимпиадные аудитории).

Скрипт выполняет полный цикл: разметка диска → установка Windows → тихая установка ПО → Windows Update → активация.

---

## Содержание

- [Что делает скрипт](#что-делает-скрипт)
- [Структура репозитория](#структура-репозитория)
- [Требования](#требования)
- [Подготовка USB-накопителя](#подготовка-usb-накопителя)
- [Настройка WiFi](#настройка-wifi)
- [Установка Windows (autounattend.xml)](#установка-windows-autounattendxml)
- [Запуск после установки](#запуск-после-установки)
- [Устанавливаемое ПО](#устанавливаемое-по)
- [Учётные записи](#учётные-записи)
- [После завершения](#после-завершения)
- [Логи и отчёт](#логи-и-отчёт)
- [Частые проблемы](#частые-проблемы)

---

## Что делает скрипт

1. **Разметка диска** — GPT, EFI 300 МБ, MSR 16 МБ, Windows 120 ГБ
2. **Установка Windows 11 Pro** (ru-RU) без диалогов и учётной записи Microsoft
3. **Создание пользователей** — `EXAM` (администратор) и `User` (без пароля)
4. **Тихая установка ПО** из папки `soft/` на USB
5. **Windows Update** — до 5 циклов обновлений без перезагрузки
6. **Подключение к WiFi** или пропуск при наличии проводной сети
7. **Отчёт на рабочем столе** — список установленного ПО и статус активации
8. **Активация Windows и Office** — ярлык на рабочем столе
9. **Перезагрузка** через 30 секунд после завершения

---

## Структура репозитория

```
/
├── autounattend.xml          # Ответный файл для автоустановки Windows
├── start.cmd                 # Запуск superscript_modified.ps1 вручную
├── wifi.txt                  # Список WiFi сетей (SSID;пароль)
│
├── soft/                     # ← Папка с дистрибутивами (НЕ в репозитории, на USB)
│   ├── superscript_modified.ps1   # Главный скрипт установки
│   ├── wing.exe
│   ├── sublime_text_build_*_x64_setup.exe
│   ├── VSCodiumSetup-x64-*.exe
│   ├── pycharm-community-*.exe
│   ├── ideaIC-*.exe
│   ├── python-3.12.4-amd64.exe
│   ├── python-3.8.10-amd64.exe
│   ├── jdk-11*_windows-x64_bin.exe
│   ├── OpenJDK21U-jdk_x64_windows_hotspot_*.msi
│   ├── 7z*-x64.exe
│   ├── tcmd*x64.exe
│   ├── Far30*.x86.*.msi
│   ├── LibreOffice_*_Win_x86-64.msi
│   ├── MicrosoftOffice2019/   # папка с Setup.exe + configuration.xml
│   ├── eclipse-java-*-win32-x86_64.zip
│   ├── codeblocks-*mingw-setup.exe
│   ├── PascalABCNETSetup.exe
│   ├── kumir2-*.exe           # или kumir-setup.exe
│   └── vs2022/
│       └── vs_setup.exe
│
└── C:\Windows\Setup\Scripts\ # ← Копируется autounattend.xml автоматически
    ├── SuperScript.ps1        # Лончер (ищет soft/ на любом диске)
    ├── FirstLogon.ps1
    ├── Specialize.ps1
    ├── DefaultUser.ps1
    └── ...
```

> **Папка `soft/` не входит в репозиторий** — её нужно наполнить дистрибутивами вручную (см. ниже).

---

## Требования

| Компонент | Минимум |
|-----------|---------|
| USB-накопитель | 32 ГБ (для Windows ISO + soft/) |
| Архитектура | x86-64 (amd64) |
| Диск компьютера | 128 ГБ SSD/HDD |
| Сеть | Опционально (WiFi или Ethernet) |
| Windows ISO | Windows 11 Pro (ru-RU) |

---

## Подготовка USB-накопителя

### 1. Создать загрузочный USB с Windows

Используйте [Rufus](https://rufus.ie/) или [Ventoy](https://www.ventoy.net/):

```
Rufus → выберите ISO → схема GPT → целевая система UEFI → старт
```

### 2. Скопировать файлы автоответа

После создания загрузочного USB скопируйте в **корень USB**:

```
autounattend.xml
wifi.txt
```

Также скопируйте в `sources\$OEM$\$$\Setup\Scripts\` на USB:

```
SuperScript.ps1
FirstLogon.ps1
Specialize.ps1
DefaultUser.ps1
RemoveCapabilities.ps1
RemoveFeatures.ps1
RemovePackages.ps1
SetStartPins.ps1
SetWallpaper.ps1
UserOnce.ps1
```

> Если папки не существует — создайте её вручную.

### 3. Создать папку `soft/` с дистрибутивами

В **корне USB** создайте папку `soft` и положите туда все установщики (см. [список](#устанавливаемое-по)).

```
USB:\
├── autounattend.xml
├── wifi.txt
├── start.cmd
└── soft\
    ├── superscript_modified.ps1
    ├── wing.exe
    ├── python-3.12.4-amd64.exe
    └── ...
```

---

## Настройка WiFi

Отредактируйте файл `wifi.txt` — по одной сети на строку:

```
# Формат: SSID;Пароль
# Строки с # игнорируются
# Скрипт пробует сети по порядку

ИмяСети1;пароль123
ИмяСети2;другойпароль
```

Скрипт **пропускает WiFi**, если обнаружено активное проводное соединение (Ethernet).  
Если `wifi.txt` отсутствует — установка продолжается без интернета (Windows Update пропускается).

---

## Установка Windows (autounattend.xml)

### Процесс

1. Вставьте USB в компьютер и загрузитесь с него (UEFI Boot)
2. Windows начнёт установку **полностью автоматически**:
   - Язык: **русский**, раскладки: RU + EN
   - Диск 0 будет **полностью очищен и переразмечен**
   - Раздел Windows: **120 ГБ**
3. После установки Windows автоматически:
   - Запустит `FirstLogon.ps1`, который вызовет `SuperScript.ps1`
   - Скрипт найдёт папку `soft/` на USB и начнёт установку ПО

> ⚠️ **Внимание:** `autounattend.xml` очищает диск 0 без предупреждений. Убедитесь, что на нём нет нужных данных.

### Обход требований Windows 11

Файл автоматически обходит проверки:
- TPM 2.0
- Secure Boot  
- RAM (минимум 4 ГБ)

---

## Запуск после установки

Если скрипт не запустился автоматически (или нужно повторить установку ПО):

1. Подключите USB с папкой `soft/`
2. Откройте `start.cmd` от имени администратора **или** запустите вручную:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "D:\soft\superscript_modified.ps1"
```

Скрипт сам найдёт папку `soft/` на любом подключённом диске.

---

## Устанавливаемое ПО

### Среды разработки и IDE

| Программа | Файл установщика |
|-----------|-----------------|
| WingIDE 101 | `wing.exe` |
| Sublime Text | `sublime_text_build_4169_x64_setup.exe` |
| VSCodium | `VSCodiumSetup-x64-1.112.01907.exe` |
| PyCharm Community | `pycharm-community-2025.1.1.1.exe` |
| IntelliJ IDEA Community | `ideaIC-2025.1.1.1.exe` |
| Code::Blocks (с MinGW) | `codeblocks-20.03mingw-setup.exe` |
| PascalABC.NET | `PascalABCNETSetup.exe` |
| KuMir | `kumir2-2.1.0-rc11-install.exe` или `kumir-setup.exe` |
| Visual Studio 2022 Community | `vs2022\vs_setup.exe` |
| Eclipse IDE (Java) | `eclipse-java-2020-09-R-win32-x86_64.zip` |

### Языки программирования

| Программа | Файл установщика |
|-----------|-----------------|
| Python 3.12.4 | `python-3.12.4-amd64.exe` |
| Python 3.8.10 | `python-3.8.10-amd64.exe` |
| JDK 11.0.21 | `jdk-11.0.21_windows-x64_bin.exe` |
| OpenJDK 21 | `OpenJDK21U-jdk_x64_windows_hotspot_21.0.7_6.msi` |

### Утилиты и офис

| Программа | Файл установщика |
|-----------|-----------------|
| 7-Zip | `7z2406-x64.exe` |
| Total Commander | `tcmd1103x64.exe` |
| Far Manager | `Far30b6300.x86.20240407.msi` |
| LibreOffice 7.6 | `LibreOffice_7.6.6_Win_x86-64.msi` |
| Microsoft Office 2019 | `MicrosoftOffice2019\Setup.exe` |

> Все программы устанавливаются в **тихом режиме** (`/S`, `/qn`, `/VERYSILENT`).  
> Если программа уже установлена — она **пропускается**.

---

## Учётные записи

| Пользователь | Пароль | Группа |
|-------------|--------|--------|
| `EXAM` | `123159` | Администраторы |
| `User` | *(пусто)* | Пользователи |

Автовход настроен для `EXAM` (однократно, для запуска скриптов установки).

> Пароли можно изменить в `autounattend.xml` перед сборкой USB.

---

## После завершения

По окончании работы скрипт:

1. Создаёт **отчёт на рабочем столе** (`Installation_Report_*.txt`) — статус каждой программы
2. Кладёт на рабочий стол файлы `Activate_Windows_Office.cmd` и `activate_helper.ps1`
3. **Перезагружает компьютер** через 30 секунд

### Активация

После перезагрузки запустите с рабочего стола:

```
Activate_Windows_Office.cmd  (от имени администратора)
```

Скрипт использует [MAS (Microsoft Activation Scripts)](https://github.com/massgravel/Microsoft-Activation-Scripts).

---

## Логи и отчёт

| Файл | Расположение |
|------|-------------|
| Лог суперскрипта | `%TEMP%\SuperScript_YYYYMMDD_HHmmss.log` |
| Лог первого входа | `C:\Windows\Setup\Scripts\FirstLogon.log` |
| Отчёт об установке | Рабочий стол → `Installation_Report_*.txt` |

---

## Частые проблемы

**Скрипт не находит папку `soft/`**  
→ Убедитесь, что USB подключён и папка называется именно `soft` (строчными буквами).

**Программа не устанавливается (файл не найден)**  
→ Проверьте, что имя файла установщика в `soft/` точно совпадает с тем, что указано в скрипте.

**WiFi не подключается**  
→ Убедитесь, что в `wifi.txt` правильный SSID и пароль. Пустые строки и строки с `#` игнорируются.

**Windows Update не запускается**  
→ Скрипт требует интернет. Без него Windows Update пропускается автоматически.

**Зависание на этапе Visual Studio**  
→ VS 2022 устанавливается долго (10–30 минут). Это нормально — не закрывайте окно.

**Диск не размечается**  
→ Убедитесь, что загрузка идёт в режиме UEFI (не Legacy/CSM). В BIOS отключите CSM.

---

## Лицензия

Проект создан для внутреннего использования в учебных заведениях.  
Скрипт активации Windows/Office использует [MAS](https://github.com/massgravel/Microsoft-Activation-Scripts) — ознакомьтесь с условиями использования.
