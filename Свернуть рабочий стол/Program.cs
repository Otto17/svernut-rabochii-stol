// Copyright (c) 2025 Otto
// Лицензия: MIT (см. LICENSE)

using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;

namespace Вернуть_рабочий_стол
{
    class Program
    {
        private static readonly string ConfigPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings.conf");   // Путь до конфига
        private static readonly string LogPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "log_Desktop.txt");    // Путь к лог-файлу
        private static readonly object LogLock = new();         // Блокировка логов
        private static readonly object SettingsLock = new();    // Блокировка настроек

        private static volatile bool _keepRunning = true; // Флаг работы фонового цикла

        // Текущее состояние настроек (обновляется из консоли)
        private static Settings _settings;

        // --- Трей и UI ---
        private static NotifyIcon _trayIcon;
        private static ContextMenuStrip _trayMenu;
        private static ToolStripMenuItem _miOpenConsole;
        private static Control _uiInvoker; // универсальный инвокер для UI-потока

        // Динамический пункт меню
        private static ToolStripMenuItem _miImageToggle;
        private static volatile bool _imageModeEnabled = false; // Режим "Картинка включена/выключена"
        private static volatile bool _isMinimizedState = false; // Находимся ли сейчас в состоянии "свернуто"

        // Оверлеи (картинка/чёрный фон)
        private static readonly object _overlayLock = new();
        private static readonly List<Form> _imageOverlays = [];
        private static Image _currentOverlayImage = null;

        private static Thread _workerThread;
        private static volatile bool _consoleOpen = false; // Открыта ли сейчас "консоль настроек"

        // Одиночный экземпляр приложения
        private static Mutex _singleInstanceMutex; // Хранит мьютекс на время жизни процесса

        // Иконки для пунктов меню (кэш)
        private static Image _iconConsole, _iconPicture, _iconLog, _iconConfig, _iconFolder, _iconExit;

        // WinAPI: управление консолью
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool AllocConsole();
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool FreeConsole();
        [DllImport("kernel32.dll")] private static extern IntPtr GetConsoleWindow();
        [DllImport("kernel32.dll")] private static extern bool SetConsoleCP(uint wCodePageID);
        [DllImport("kernel32.dll")] private static extern bool SetConsoleOutputCP(uint wCodePageID);

        // WinAPI: низкоуровневый ввод/вывод Unicode
        [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr GetStdHandle(int nStdHandle);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool ReadConsoleW(IntPtr hConsoleInput, [Out] char[] lpBuffer, uint nNumberOfCharsToRead, out uint lpNumberOfCharsRead, IntPtr pInputControl);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool WriteConsoleW(IntPtr hConsoleOutput, string lpBuffer, uint nNumberOfCharsToWrite, out uint lpNumberOfCharsWritten, IntPtr lpReserved);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);

        private const int STD_INPUT_HANDLE = -10;
        private const int STD_OUTPUT_HANDLE = -11;
        private const int STD_ERROR_HANDLE = -12;

        private const uint ENABLE_PROCESSED_INPUT = 0x0001;
        private const uint ENABLE_LINE_INPUT = 0x0002;
        private const uint ENABLE_ECHO_INPUT = 0x0004;

        // WinAPI: сворачивание/разворачивание (через Shell_TrayWnd)
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        // WinAPI: резервный вариант — эмуляция Win + D
        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        [StructLayout(LayoutKind.Sequential)]
        private struct INPUT
        {
            public uint type;
            public InputUnion U;
        }

        [StructLayout(LayoutKind.Explicit)]
        private struct InputUnion
        {
            [FieldOffset(0)]
            public KEYBDINPUT ki;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        private const uint INPUT_KEYBOARD = 1;
        private const uint KEYEVENTF_KEYUP = 0x0002;
        private const ushort VK_LWIN = 0x5B;
        private const ushort VK_D = 0x44;

        private const uint WM_COMMAND = 0x0111;
        private static readonly IntPtr MIN_ALL = (IntPtr)419;       // Показать рабочий стол (свернуть все)
        private static readonly IntPtr MIN_ALL_UNDO = (IntPtr)416;  // Отменить сворачивание

        // Дескрипторы стандартных потоков консоли
        private static IntPtr _hStdin = IntPtr.Zero;
        private static IntPtr _hStdout = IntPtr.Zero;
        private static IntPtr _hStderr = IntPtr.Zero;

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Одиночный экземпляр: если уже запущено — покажем сообщение и выйдет
            _singleInstanceMutex = new Mutex(true, @"Local\ВернутьРабочийСтол_SingleInstance", out bool isNew);
            if (!isNew)
            {
                MessageBox.Show("Программа уже запущена в трее", "Вернуть рабочий стол", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Инициализация лог-файла (при первом запуске — авторство)
            try
            {
                if (!File.Exists(LogPath))
                {
                    LogAuthorInfo();
                }
            }
            catch { /* Игнорируем */ }

            try
            {
                // Чтение настроек (могут отсутствовать на первом запуске)
                _settings = LoadSettings();
                if (_settings == null)
                {
                    AppendToLog("Конфиг не найден или повреждён. Настройте через «Консоль настроек…» в меню трея.");
                }
                else
                {
                    _imageModeEnabled = _settings.ImageModeEnabled; // восстановим состояние картинки
                }

                // Создадим иконку трея и меню и UI-инвокер (до старта воркера!)
                CreateTray();

                // Подсказка при первом старте или при отсутствии конфига
                try
                {
                    _trayIcon.BalloonTipTitle = "Вернуть рабочий стол";
                    _trayIcon.BalloonTipText = _settings == null
                    ? "Нужна настройка. Откройте «Консоль настроек…» из меню трея."
                    : "Приложение работает в области уведомлений.";
                    _trayIcon.ShowBalloonTip(3500);
                }
                catch { /* Не критично */ }

                // Запускаем фоновую работу после подготовки UI
                _workerThread = new Thread(WorkerLoop) { IsBackground = true };
                _workerThread.Start();

                // Запуск message loop
                Application.Run();

                // Завершение: остановим фон и подчистим трей
                _keepRunning = false;
                try { _workerThread?.Join(2000); } catch { }
                try
                {
                    if (_trayIcon != null)
                    {
                        _trayIcon.Visible = false;
                        _trayIcon.Dispose();
                    }
                }
                catch { }
            }
            catch (Exception ex)
            {
                AppendToLog("Критическая ошибка: " + ex.Message);
                MessageBox.Show("Критическая ошибка:\n" + ex.Message, "Вернуть рабочий стол", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Фоновый рабочий цикл: сворачивает/восстанавливает окна по расписанию
        private static void WorkerLoop()
        {
            bool? minimizedNow = null;

            while (_keepRunning)
            {
                try
                {
                    Settings s = null;
                    lock (SettingsLock) s = _settings;

                    if (s == null)
                    {
                        Thread.Sleep(1000);
                        continue;
                    }

                    var now = DateTime.Now.TimeOfDay;
                    bool shouldMinimize = IsTimeInRange(now, s.StartTime, s.EndTime);

                    if (!minimizedNow.HasValue || shouldMinimize != minimizedNow.Value)
                    {
                        if (shouldMinimize)
                        {
                            if (MinimizeAllWindows(out _))
                                AppendToLog($"Все окна свёрнуты. Причина: {(minimizedNow.HasValue ? "смена периода" : "старт работы")}.");
                            else
                                AppendToLog($"Не удалось свернуть окна (попытка).");

                            // Показ картинки/чёрного фона (если режим включён)
                            if (_imageModeEnabled)
                            {
                                ShowImageOverlay();
                            }
                        }
                        else
                        {
                            // Сначала закрываем картинку/фон, затем разворачиваем окна
                            CloseImageOverlay();

                            if (!minimizedNow.HasValue)
                            {
                                if (RestoreAllWindows(out _)) { /* без подробностей */ }
                                AppendToLog("Программа запущена");
                            }
                            else
                            {
                                if (RestoreAllWindows(out string method))
                                    AppendToLog($"Восстановили окна (метод: {method}). Причина: смена периода.");
                                else
                                    AppendToLog($"Не удалось восстановить окна (попытка).");
                            }
                        }

                        minimizedNow = shouldMinimize;
                        _isMinimizedState = shouldMinimize;
                    }

                    Thread.Sleep(1000);
                }
                catch (Exception ex)
                {
                    AppendToLog($"[Worker] Ошибка: {ex.Message}");
                    Thread.Sleep(3000);
                }
            }

            AppendToLog("Завершение работы...");
        }

        // Создание иконки трея и меню
        private static void CreateTray()
        {
            _trayIcon = new NotifyIcon
            {
                Icon = LoadTrayIcon(),
                Visible = true,
                Text = "Вернуть рабочий стол"
            };

            // Меню с иконками слева
            _trayMenu = new ContextMenuStrip
            {
                ShowImageMargin = true
            };

            // Подгружаем PNG-иконки из Embedded Resource (папка Resources)
            _iconConsole = LoadMenuIcon("ConsoleIcon.png");
            _iconPicture = LoadMenuIcon("PictureIcon.png");
            _iconLog = LoadMenuIcon("LogIcon.png");
            _iconConfig = LoadMenuIcon("ConfigIcon.png");
            _iconFolder = LoadMenuIcon("FolderIcon.png");
            _iconExit = LoadMenuIcon("exitIcon.png");

            // Создаём универсальный инвокер на UI-потоке заранее
            _uiInvoker = new Control();
            _uiInvoker.CreateControl();

            // Пункты меню с иконками
            _miOpenConsole = new ToolStripMenuItem("Консоль настроек...", _iconConsole, (s, e) => OpenSettingsConsoleAsync());
            _trayMenu.Items.Add(_miOpenConsole);

            // Динамический пункт (иконка одна и та же)
            _miImageToggle = new ToolStripMenuItem("Картинка выключена", _iconPicture, (s, e) => ToggleImageMode());
            _trayMenu.Items.Add(_miImageToggle);
            UpdateImageMenuText();

            _trayMenu.Items.Add(new ToolStripSeparator());
            _trayMenu.Items.Add(new ToolStripMenuItem("Открыть лог", _iconLog, (s, e) => OpenWithNotepad(LogPath)));
            _trayMenu.Items.Add(new ToolStripMenuItem("Открыть конфиг", _iconConfig, (s, e) => OpenWithNotepad(ConfigPath)));
            _trayMenu.Items.Add(new ToolStripMenuItem("Папка программы", _iconFolder, (s, e) => OpenFolder(AppDomain.CurrentDomain.BaseDirectory)));

            _trayMenu.Items.Add(new ToolStripSeparator());
            _trayMenu.Items.Add(new ToolStripMenuItem("Выход", _iconExit, (s, e) => ExitApp()));

            _trayIcon.ContextMenuStrip = _trayMenu;
            _trayIcon.DoubleClick += (s, e) => OpenSettingsConsoleAsync();
            Application.ApplicationExit += (s, e) =>
            {
                try
                {
                    if (_trayIcon != null)
                    {
                        _trayIcon.Visible = false;
                        _trayIcon.Dispose();
                    }
                    _uiInvoker?.Dispose();
                }
                catch { }
            };
        }

        // Переключение режима "Картинка включена/выключена"
        private static void ToggleImageMode()
        {
            _imageModeEnabled = !_imageModeEnabled;
            UpdateImageMenuText();
            AppendToLog(_imageModeEnabled ? "Режим «Картинка включена» активирован." : "Режим «Картинка выключена» активирован.");

            // Сохраняем в конфиг
            lock (SettingsLock)
            {
                if (_settings != null)
                {
                    _settings.ImageModeEnabled = _imageModeEnabled;
                    SaveSettings(_settings, "обновлены", writeLog: false); // не логируем блок "Настройки обновлены"
                }
            }

            // Если уже в состоянии "свернуто", применяем немедленно
            if (_imageModeEnabled && _isMinimizedState)
            {
                ShowImageOverlay();
            }
            else
            {
                CloseImageOverlay();
            }
        }

        // Обновляет текст в трее
        private static void UpdateImageMenuText()
        {
            try
            {
                if (_miImageToggle != null && !_miImageToggle.IsDisposed)
                {
                    _miImageToggle.Text = _imageModeEnabled ? "Картинка включена" : "Картинка выключена";
                }
            }
            catch { }
        }

        // Показ картинки/чёрного фона на всех экранах (на UI потоке)
        private static void ShowImageOverlay()
        {
            RunOnUiThread(() =>
            {
                lock (_overlayLock)
                {
                    if (_imageOverlays.Count > 0) return; // Уже показано
                }

                string imgPath = GetFirstImagePathOrNull();
                Image img = null;

                if (imgPath != null)
                {
                    try
                    {
                        using var fs = new FileStream(imgPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                        img = Image.FromStream(fs); // Загружает в память, чтобы не блокировать файл
                    }
                    catch (Exception ex)
                    {
                        AppendToLog($"Не удалось загрузить картинку '{imgPath}': {ex.Message}. Будет показан чёрный фон.");
                        img = null;
                    }
                }

                foreach (var screen in Screen.AllScreens)
                {
                    var form = new Form
                    {
                        FormBorderStyle = FormBorderStyle.None,
                        StartPosition = FormStartPosition.Manual,
                        Bounds = screen.Bounds,
                        BackColor = Color.Black,
                        TopMost = false,    // Не topmost — не блокирует Диспетчер задач и Win
                        ShowInTaskbar = false,
                        MinimizeBox = false,
                        MaximizeBox = false
                    };

                    if (img != null)
                    {
                        var pb = new PictureBox
                        {
                            Dock = DockStyle.Fill,
                            SizeMode = PictureBoxSizeMode.Zoom,
                            Image = img
                        };
                        form.Controls.Add(pb);
                    }

                    form.FormClosed += (s, e) =>
                    {
                        lock (_overlayLock)
                        {
                            _imageOverlays.Remove(form);
                        }
                    };

                    lock (_overlayLock)
                    {
                        _imageOverlays.Add(form);
                    }

                    try { form.Show(); } catch { }
                }

                // Запоминаем картинку, чтобы потом освободить
                _currentOverlayImage = img;

                AppendToLog(img != null
                ? $"Открыта картинка на всех экранах: {Path.GetFileName(imgPath)}"
                : "Открыт чёрный фон на всех экранах.");
            });
        }

        // Закрытие картинок/фона (на UI потоке)
        private static void CloseImageOverlay()
        {
            RunOnUiThread(() =>
            {
                lock (_overlayLock)
                {
                    if (_imageOverlays.Count == 0 && _currentOverlayImage == null) return;

                    foreach (var f in _imageOverlays.ToList())
                    {
                        try { f.Close(); f.Dispose(); } catch { }
                    }
                    _imageOverlays.Clear();

                    try { _currentOverlayImage?.Dispose(); } catch { }
                    _currentOverlayImage = null;
                }

                AppendToLog("Картинка/фон закрыты.");
            });
        }

        // Асинхронный запуск консоли настроек (не блокируем UI)
        private static void OpenSettingsConsoleAsync()
        {
            if (_consoleOpen) return;
            _consoleOpen = true;
            if (_miOpenConsole != null) _miOpenConsole.Enabled = false;

            var t = new Thread(() =>
            {
                try
                {
                    RunSettingsConsole();
                }
                finally
                {
                    try
                    {
                        if (_trayMenu != null && !_trayMenu.IsDisposed)
                        {
                            _trayMenu.BeginInvoke((Action)(() =>
                            {
                                if (_miOpenConsole != null) _miOpenConsole.Enabled = true;
                            }));
                        }
                    }
                    catch { }
                    _consoleOpen = false;
                }
            })
            {
                IsBackground = true
            };
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
        }

        // Содержимое интерактивной "консоли настроек"
        private static void RunSettingsConsole()
        {
            try
            {
                if (!AllocConsole())
                {
                    AppendToLog("Не удалось создать консоль настроек.");
                    MessageBox.Show("Не удалось открыть консоль настроек.", "Вернуть рабочий стол", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Привязываем потоки и настраиваем кодировки
                BindConsoleStreams();

                // Блокируем закрытие процесса по "Ctrl+C" в консоли
                Console.CancelKeyPress += (s, e) => { e.Cancel = true; };

                ShowAuthorInfoConsole();

                Settings current;
                lock (SettingsLock) current = _settings;

                if (current != null)
                {
                    Console.WriteLine("Текущие настройки:");
                    PrintSettings(current);
                    Console.Write("\nИзменить настройки? (Д/Н): ");
                    var answer = ReadLineFromConsole(); // Unicode-чтение

                    if (IsYesAnswer(answer))
                    {
                        var updated = AskSettingsFromUser();
                        // Сохраняем текущий режим картинки
                        updated.ImageModeEnabled = current.ImageModeEnabled;
                        SaveSettings(updated, "обновлены");
                        lock (SettingsLock)
                        {
                            _settings = updated;
                            _imageModeEnabled = _settings.ImageModeEnabled;
                        }
                        UpdateImageMenuText();
                        Console.WriteLine("\nНастройки обновлены.");
                    }
                    else
                    {
                        Console.WriteLine("Оставляем текущие настройки без изменений.");
                    }
                }
                else
                {
                    Console.WriteLine("Конфигурация отсутствует или повреждена — требуется первичная настройка.\n");
                    var created = AskSettingsFromUser();
                    created.ImageModeEnabled = _imageModeEnabled; // Берём текущее состояние пункта меню
                    SaveSettings(created, "созданы");
                    lock (SettingsLock) _settings = created;
                    Console.WriteLine("\nНастройки сохранены.");
                }

                Console.WriteLine("\nГотово. Консоль будет закрыта.");
                Thread.Sleep(500); // Короткая пауза, чтобы вывести текст
            }
            catch (Exception ex)
            {
                AppendToLog("Ошибка в консоли настроек: " + ex.Message);
                try { Console.WriteLine("Ошибка: " + ex.Message); Thread.Sleep(1200); } catch { }
            }
            finally
            {
                try { UnbindConsoleStreams(); } catch { }
                try { FreeConsole(); } catch { }
            }
        }

        private static void ExitApp()
        {
            // Если открыта консоль настроек — предупредим
            if (_consoleOpen)
            {
                var res = MessageBox.Show("Сначала закройте «Консоль настроек». Закрыть принудительно?",
                "Вернуть рабочий стол", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (res == DialogResult.Yes)
                {
                    try { FreeConsole(); } catch { }
                }
                else
                {
                    return;
                }
            }

            _keepRunning = false;
            try { _workerThread?.Join(1500); } catch { }

            // Закрыть оверлеи перед выходом
            try { CloseImageOverlay(); } catch { }

            try
            {
                if (_trayIcon != null)
                {
                    _trayIcon.Visible = false;
                    _trayIcon.Dispose();
                }
            }
            catch { }

            // Освобождаем мьютекс одиночного экземпляра
            try { _singleInstanceMutex?.ReleaseMutex(); _singleInstanceMutex?.Dispose(); } catch { }

            Application.Exit();
        }

        // Консоль: привязка и отвязка потоков ввода-вывода
        private static void BindConsoleStreams()
        {
            try
            {
                ConfigureConsoleEncoding(); // Настроим кодировку

                // Получаем стандартные дескрипторы
                _hStdin = GetStdHandle(STD_INPUT_HANDLE);
                _hStdout = GetStdHandle(STD_OUTPUT_HANDLE);
                _hStderr = GetStdHandle(STD_ERROR_HANDLE);

                // ВЫВОД: перенаправляем на Unicode-писатели (WriteConsoleW).
                var stdOutStream = Console.OpenStandardOutput();
                var stdErrStream = Console.OpenStandardError();
                Console.SetOut(new UnicodeConsoleWriter(_hStdout, stdOutStream));
                Console.SetError(new UnicodeConsoleWriter(_hStderr, stdErrStream));

                // ВВОД: оставляем поток, но для подтверждения читаем через ReadConsoleW
                var inStream = new StreamReader(Console.OpenStandardInput(), Console.InputEncoding);
                Console.SetIn(inStream);

                try { Console.Title = "Вернуть рабочий стол — консоль настроек"; } catch { }
            }
            catch { /* Если окружение не позволяет — просто молча продолжаем */ }
        }

        // Пытаемся настроить консоль на устойчивую кириллицу: CP866 -> CP1251 -> UTF-8
        private static void ConfigureConsoleEncoding()
        {
            try
            {
                // OEM 866 работает с кириллицей без кракозябр
                if (SetConsoleCP(866) && SetConsoleOutputCP(866))
                {
                    var e = Encoding.GetEncoding(866);
                    Console.InputEncoding = e;
                    Console.OutputEncoding = e;
                    return;
                }

                // Windows-1251
                if (SetConsoleCP(1251) && SetConsoleOutputCP(1251))
                {
                    var e = Encoding.GetEncoding(1251);
                    Console.InputEncoding = e;
                    Console.OutputEncoding = e;
                    return;
                }

                // UTF-8 — запасной вариант
                if (SetConsoleCP(65001) && SetConsoleOutputCP(65001))
                {
                    var e = new UTF8Encoding(false);
                    Console.InputEncoding = e;
                    Console.OutputEncoding = e;
                    return;
                }

                // В крайнем случае — штатные API
                Console.InputEncoding = Encoding.UTF8;
                Console.OutputEncoding = Encoding.UTF8;
            }
            catch
            {
                try
                {
                    Console.InputEncoding = Encoding.UTF8;
                    Console.OutputEncoding = Encoding.UTF8;
                }
                catch { }
            }
        }

        // "Отвязывает" текущее приложение от всех стандартных потоков консоли
        private static void UnbindConsoleStreams()
        {
            try
            {
                Console.SetOut(TextWriter.Synchronized(TextWriter.Null));
                Console.SetError(TextWriter.Synchronized(TextWriter.Null));
                Console.SetIn(new StreamReader(Stream.Null));
                _hStdin = _hStdout = _hStderr = IntPtr.Zero;
            }
            catch { }
        }

        // Чтение строки Unicode напрямую через ReadConsoleW (устойчиво к раскладке/кодpage)
        private static string ReadLineFromConsole()
        {
            try
            {
                if (_hStdin == IntPtr.Zero) _hStdin = GetStdHandle(STD_INPUT_HANDLE);
                if (_hStdin == IntPtr.Zero || _hStdin == new IntPtr(-1))
                {
                    var fallback = Console.ReadLine() ?? "";
                    return fallback;
                }

                // Обеспечим линейный/эхо режим (по возможности)
                if (GetConsoleMode(_hStdin, out uint mode))
                {
                    uint desired = mode | ENABLE_LINE_INPUT | ENABLE_ECHO_INPUT | ENABLE_PROCESSED_INPUT;
                    if (desired != mode) SetConsoleMode(_hStdin, desired);
                }

                char[] buf = new char[1024];
                if (ReadConsoleW(_hStdin, buf, (uint)buf.Length, out uint read, IntPtr.Zero))
                {
                    int len = (int)read;
                    if (len <= 0) return string.Empty;
                    string s = new(buf, 0, len);

                    // Убираем \r\n в конце
                    s = s.Replace("\r", "").Replace("\n", "");
                    return s;
                }

                // Фолбэк
                return Console.ReadLine() ?? "";
            }
            catch
            {
                return Console.ReadLine() ?? "";
            }
        }

        // Вывод информации об авторе (только в консоль)
        static void ShowAuthorInfoConsole()
        {
            Console.WriteLine("Автор Otto (ver.09.08.25)");
            Console.WriteLine("Ссылка на проект: https://gitflic.ru/project/otto/svernut-rabochii-stol");
            Console.WriteLine();
        }

        // Запись информации об авторе в лог (только при создании лога)
        static void LogAuthorInfo()
        {
            AppendToLog("Автор Otto (ver.09.08.25)");
            AppendToLog("Ссылка на проект: https://gitflic.ru/project/otto/svernut-rabochii-stol\n");
        }

        // Печать текущих настроек (в консоль)
        static void PrintSettings(Settings s)
        {
            string period = string.Format("{0} — {1}", FormatTime(s.StartTime), FormatTime(s.EndTime));
            bool crossesMidnight = s.StartTime > s.EndTime;

            Console.WriteLine("- Действие в период: сворачивать все окна (показать рабочий стол)");
            Console.WriteLine("- Действие вне периода: восстанавливать окна");
            Console.WriteLine("- Период: {0}{1}", period, (crossesMidnight ? " (с переходом через полночь)" : ""));
            Console.WriteLine("- Режим картинки: {0}", s.ImageModeEnabled ? "включена" : "выключена");
        }

        // Запрос настроек пользователя (консоль)
        static Settings AskSettingsFromUser()
        {
            var s = new Settings();

            // Время начала и конца периода
            Console.Write("Введите время начала периода (чч:мм:сс или чч:мм): ");
            s.StartTime = ReadTime();

            Console.Write("Введите время окончания периода (чч:мм:сс или чч:мм): ");
            s.EndTime = ReadTime();

            // Предупреждение о круглосуточном режиме
            if (s.StartTime == s.EndTime)
                Console.WriteLine("Предупреждение: время начала равно времени окончания — период будет круглосуточным");

            return s;
        }

        // Читает время из консоли и парсит его в TimeSpan
        static TimeSpan ReadTime()
        {
            while (true)
            {
                var input = (Console.ReadLine() ?? "").Trim();
                if (TryParseTime(input, out TimeSpan ts)) return ts;
                Console.Write("Неверный формат. Введите время как чч:мм:сс или чч:мм: ");
            }
        }

        // Пробуем разобрать строку в TimeSpan с поддержкой разных форматов
        static bool TryParseTime(string s, out TimeSpan ts)
        {
            string[] formats = [@"hh\:mm\:ss", @"h\:m\:s", @"hh\:mm", @"h\:m", @"h\:mm", @"hh\:m"];
            return TimeSpan.TryParseExact(s, formats, CultureInfo.InvariantCulture, out ts);
        }

        // Форматируем время в строку hh:mm:ss
        static string FormatTime(TimeSpan ts) { return ts.ToString(@"hh\:mm\:ss"); }

        // Проверка, входит ли текущее время в заданный диапазон
        static bool IsTimeInRange(TimeSpan now, TimeSpan start, TimeSpan end)
        {
            if (start == end) return true;
            if (start < end) return (now >= start && now < end);
            else return (now >= start || now < end);
        }

        // Сворачивает все окна
        private static bool MinimizeAllWindows(out string method)
        {
            method = "";
            try
            {
                var tray = FindWindow("Shell_TrayWnd", null);
                if (tray != IntPtr.Zero)
                {
                    SendMessage(tray, WM_COMMAND, MIN_ALL, IntPtr.Zero);
                    method = "WM_COMMAND:MIN_ALL";
                    return true;
                }

                // Резервный вариант — Win + D
                if (SendWinD())
                {
                    method = "SendInput:Win+D";
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                AppendToLog("MinimizeAllWindows error: " + ex.Message);
                return false;
            }
        }

        // Восстанавливает все окна
        private static bool RestoreAllWindows(out string method)
        {
            method = "";
            try
            {
                var tray = FindWindow("Shell_TrayWnd", null);
                if (tray != IntPtr.Zero)
                {
                    SendMessage(tray, WM_COMMAND, MIN_ALL_UNDO, IntPtr.Zero);
                    method = "WM_COMMAND:MIN_ALL_UNDO";
                    return true;
                }

                // Резервный вариант — Win + D (тоггл)
                if (SendWinD())
                {
                    method = "SendInput:Win+D";
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                AppendToLog("RestoreAllWindows error: " + ex.Message);
                return false;
            }
        }

        // Эмуляция нажатия Win + D через SendInput
        private static bool SendWinD()
        {
            try
            {
                var inputs = new List<INPUT>
                {
                new() { type = INPUT_KEYBOARD, U = new InputUnion { ki = new KEYBDINPUT { wVk = VK_LWIN, wScan = 0, dwFlags = 0, time = 0, dwExtraInfo = IntPtr.Zero } } },
                new() { type = INPUT_KEYBOARD, U = new InputUnion { ki = new KEYBDINPUT { wVk = VK_D, wScan = 0, dwFlags = 0, time = 0, dwExtraInfo = IntPtr.Zero } } },
                new() { type = INPUT_KEYBOARD, U = new InputUnion { ki = new KEYBDINPUT { wVk = VK_D, wScan = 0, dwFlags = KEYEVENTF_KEYUP, time = 0, dwExtraInfo = IntPtr.Zero } } },
                new() { type = INPUT_KEYBOARD, U = new InputUnion { ki = new KEYBDINPUT { wVk = VK_LWIN, wScan = 0, dwFlags = KEYEVENTF_KEYUP, time = 0, dwExtraInfo = IntPtr.Zero } } }
                };

                uint sent = SendInput((uint)inputs.Count, [.. inputs], Marshal.SizeOf(typeof(INPUT)));
                return sent == inputs.Count;
            }
            catch
            {
                return false;
            }
        }

        // Читает конфиг-файл, возвращает настройки или null, если он битый/отсутствует
        static Settings LoadSettings()
        {
            try
            {
                if (!File.Exists(ConfigPath)) return null;
                var lines = File.ReadAllLines(ConfigPath, Encoding.UTF8);

                // Формат: key=value, допускаются комментарии (#) и произвольный порядок
                var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var raw in lines)
                {
                    var line = (raw ?? "").Trim();
                    if (line.Length == 0) continue;
                    if (line.StartsWith("#")) continue;
                    int eq = line.IndexOf('=');
                    if (eq <= 0) continue;
                    var key = line.Substring(0, eq).Trim();
                    var val = line.Substring(eq + 1).Trim();
                    if (key.Length > 0) dict[key] = val;
                }
                if (dict.Count > 0)
                {
                    if (dict.TryGetValue("start_time", out var s1) &&
                    dict.TryGetValue("end_time", out var s2) &&
                    TryParseTime(s1, out var start) &&
                    TryParseTime(s2, out var end))
                    {
                        bool img = false;
                        if (dict.TryGetValue("image_display", out var disp))
                            TryParseBoolConfig(disp, out img);
                        return new Settings { StartTime = start, EndTime = end, ImageModeEnabled = img };
                    }
                }

                // Совместимость: старые форматы (2 строки времени; либо 4 строки с временем в 3–4)
                if (lines.Length >= 2)
                {
                    if (TryParseTime((lines[0] ?? "").Trim(), out TimeSpan start) &&
                    TryParseTime((lines[1] ?? "").Trim(), out TimeSpan end))
                    {
                        return new Settings { StartTime = start, EndTime = end, ImageModeEnabled = false };
                    }
                }
                if (lines.Length >= 4)
                {
                    if (TryParseTime((lines[2] ?? "").Trim(), out TimeSpan s) &&
                    TryParseTime((lines[3] ?? "").Trim(), out TimeSpan e))
                    {
                        return new Settings { StartTime = s, EndTime = e, ImageModeEnabled = false };
                    }
                }

                return null;
            }
            catch { return null; }
        }

        // Сохраняет настройки в файл (с комментариями/справкой); лог настроек — по флагу
        static void SaveSettings(Settings s, string action, bool writeLog = true)
        {
            try
            {
                var sb = new StringBuilder();
                sb.AppendLine("# Время начала (чч:мм:сс)");
                sb.AppendLine($"start_time={s.StartTime:hh\\:mm\\:ss}");
                sb.AppendLine();
                sb.AppendLine("# Время окончания (чч:мм:сс)");
                sb.AppendLine($"end_time={s.EndTime:hh\\:mm\\:ss}");
                sb.AppendLine();
                sb.AppendLine("# Флаг показа картинки (true - показать, false - не показывать)");
                sb.AppendLine($"image_display={(s.ImageModeEnabled ? "true" : "false")}");
                sb.AppendLine();
                sb.AppendLine();
                sb.AppendLine("# Справка");
                sb.AppendLine();
                sb.AppendLine("# Установка интервала времени поддерживается:");
                sb.AppendLine("# В прямом порядке (внутри одного календарного дня), например с 11:00:00 до 17:05:10");
                sb.AppendLine("# В инверсном виде (через полночь), например с 07:40:00 до 19:25:00");
                sb.AppendLine();
                sb.AppendLine("# Показ картинки:");
                sb.AppendLine("# Картинка берётся из этой же папки, где лежит исполняемый файл \"Свернуть рабочий стол.exe\"");
                sb.AppendLine("# Название картинки может быть любым");
                sb.AppendLine("# Поддерживаются картинки с расширениями: .jpg, .jpeg, .gif, .png");
                sb.AppendLine("# Если в папке более одной картинки, тогда берётся первая картинка в алфавитном порядке");
                sb.AppendLine("# Если в папке нет картинок, тогда показывается чёрный фон");

                File.WriteAllText(ConfigPath, sb.ToString(), Encoding.UTF8);

                if (writeLog)
                {
                    LogSettings(s, action);
                }
            }
            catch (Exception ex)
            {
                AppendToLog("Ошибка сохранения настроек: " + ex.Message);
            }
        }

        // Запись события в лог
        static void AppendToLog(string line)
        {
            try
            {
                var full = string.Format("[{0:yyyy-MM-dd HH:mm:ss}] {1}", DateTime.Now, line);
                lock (LogLock)
                {
                    File.AppendAllText(LogPath, full + Environment.NewLine, Encoding.UTF8);
                }
            }
            catch { }
        }

        // Подробная запись настроек в лог
        static void LogSettings(Settings s, string action)
        {
            try
            {
                var sb = new StringBuilder();
                sb.AppendLine();
                sb.AppendLine();
                sb.AppendLine(string.Format("===== [{0:yyyy-MM-dd HH:mm:ss}] Настройки {1} =====", DateTime.Now, action));
                bool crossesMidnight = s.StartTime > s.EndTime;
                sb.AppendLine("Действие в период: сворачивать все окна");
                sb.AppendLine("Действие вне периода: восстанавливать окна");
                sb.AppendLine(string.Format("Период: {0} — {1}{2}",
                FormatTime(s.StartTime),
                FormatTime(s.EndTime),
                (crossesMidnight ? " (с переходом через полночь)" : "")));
                sb.AppendLine("Режим картинки: " + (s.ImageModeEnabled ? "включена" : "выключена"));
                sb.AppendLine();
                lock (LogLock)
                {
                    File.AppendAllText(LogPath, sb.ToString(), Encoding.UTF8);
                }
            }
            catch { }
        }

        // Загрузка "icon.ico" для иконки трея
        private static Icon LoadTrayIcon()
        {
            try
            {
                var asm = typeof(Program).Assembly;
                foreach (var res in asm.GetManifestResourceNames())
                {
                    if (res.EndsWith(".ico", StringComparison.OrdinalIgnoreCase))
                    {
                        using var stream = asm.GetManifestResourceStream(res);
                        if (stream != null)
                            return new Icon(stream);
                    }
                }
            }
            catch { }

            // Если не нашли ресурс — стандартная иконка
            return SystemIcons.Application;
        }

        // Открыть файл в блокноте
        private static void OpenWithNotepad(string path)
        {
            try
            {
                if (!File.Exists(path))
                    File.WriteAllText(path, "", Encoding.UTF8);

                Process.Start("notepad.exe", "\"" + path + "\"");
            }
            catch (Exception ex)
            {
                AppendToLog("Ошибка открытия файла: " + ex.Message);
            }
        }

        // Открыть папку
        private static void OpenFolder(string folder)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = folder,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                AppendToLog("Ошибка открытия папки: " + ex.Message);
            }
        }

        // Вспомогательное: выполнение на UI-потоке
        private static void RunOnUiThread(Action action)
        {
            try
            {
                var ctl = _uiInvoker;
                if (ctl != null && !ctl.IsDisposed)
                {
                    if (ctl.InvokeRequired) ctl.BeginInvoke(action);
                    else action();
                }
            }
            catch { }
        }

        // Поиск первой картинки по алфавиту (или null, если нет)
        private static string GetFirstImagePathOrNull()
        {
            try
            {
                string dir = AppDomain.CurrentDomain.BaseDirectory;
                var allowedExts = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { ".jpg", ".jpeg", ".gif", ".png" };

                var files = Directory.EnumerateFiles(dir)
                .Where(f => allowedExts.Contains(Path.GetExtension(f)))
                .OrderBy(f => Path.GetFileName(f), StringComparer.CurrentCultureIgnoreCase)
                .ToList();

                return files.FirstOrDefault();
            }
            catch
            {
                return null;
            }
        }

        // Парсер логического параметра из конфига (true/false + 1/0)
        private static bool TryParseBoolConfig(string s, out bool val)
        {
            if (bool.TryParse((s ?? "").Trim(), out val)) return true;
            var t = (s ?? "").Trim();
            if (t == "1") { val = true; return true; }
            if (t == "0") { val = false; return true; }
            val = false; return false;
        }

        // Принятие подтверждения: только 'Y', 'Д', 'ДА' (независимо от регистра)
        private static bool IsYesAnswer(string s)
        {
            var t = (s ?? "").Trim().ToUpperInvariant();
            return t == "Y" || t == "Д" || t == "ДА";
        }

        // Загрузка PNG-иконки из Embedded Resource (папка Resources)
        private static Image LoadMenuIcon(string fileName)
        {
            try
            {
                var asm = typeof(Program).Assembly;
                string resName = asm.GetManifestResourceNames()
                    .FirstOrDefault(n =>
                        n.EndsWith(".Resources." + fileName, StringComparison.OrdinalIgnoreCase) ||
                        n.EndsWith("." + fileName, StringComparison.OrdinalIgnoreCase));
                if (resName == null) return null;
                using var s = asm.GetManifestResourceStream(resName);
                if (s == null) return null;
                return Image.FromStream(s);
            }
            catch { return null; }
        }

        // Unicode-писатель для консоли
        class UnicodeConsoleWriter : TextWriter
        {
            private readonly IntPtr _handle;
            private readonly TextWriter _fallback;
            private static readonly Encoding Utf8NoBom = new UTF8Encoding(false);

            public UnicodeConsoleWriter(IntPtr handle, Stream fallbackStream)
            {
                _handle = handle;
                try
                {
                    _fallback = new StreamWriter(fallbackStream, Utf8NoBom) { AutoFlush = true, NewLine = Environment.NewLine };
                }
                catch
                {
                    _fallback = TextWriter.Synchronized(TextWriter.Null);
                }
            }

            public override Encoding Encoding => Utf8NoBom;

            private bool CanUseWriteConsole =>
            _handle != IntPtr.Zero && _handle != new IntPtr(-1);

            public override void Write(char value) => Write(value.ToString());

            public override void Write(string value)
            {
                if (string.IsNullOrEmpty(value)) return;

                if (CanUseWriteConsole)
                {
                    try
                    {
                        WriteConsoleW(_handle, value, (uint)value.Length, out _, IntPtr.Zero);
                        return;
                    }
                    catch { /* Упадём в фолбэк */ }
                }
                _fallback.Write(value);
            }

            public override void WriteLine(string value)
            {
                Write(value);
                Write(Environment.NewLine);
            }

            public override void Flush()
            {
                try { _fallback.Flush(); } catch { }
            }
        }
    }

    // Класс "Settings" — хранит все пользовательские настройки программы
    class Settings
    {
        public TimeSpan StartTime { get; set; } // Время начала периода сворачивания
        public TimeSpan EndTime { get; set; } // Время окончания периода сворачивания
        public bool ImageModeEnabled { get; set; } // Состояние показа картинки
    }
}