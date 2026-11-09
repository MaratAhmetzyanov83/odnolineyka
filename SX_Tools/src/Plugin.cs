using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: ExtensionApplication(typeof(SXTools.Plugin))]
[assembly: CommandClass(typeof(SXTools.Commands))]

namespace SXTools
{
    /// <summary>Загрузка LISP-инструментов и вкладка «SX» на ленте. Корень = папка SX_Tools (над Contents).</summary>
    public class Plugin : IExtensionApplication
    {
        internal static string Root;

        public void Initialize()
        {
            string dll = Assembly.GetExecutingAssembly().Location;
            Root = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(dll), ".."));
            AcApp.Idle += OnIdle;
            AcApp.DocumentManager.DocumentCreated += (s, e) => LoadLisp(e.Document, false);
            // Смена рабочего пространства пересобирает ленту — вернуть вкладку.
            AcApp.SystemVariableChanged += (s, e) =>
            {
                if (string.Equals(e.Name, "WSCURRENT", StringComparison.OrdinalIgnoreCase)) SxRibbon.Ensure();
            };
        }

        public void Terminate() { }

        static void OnIdle(object sender, EventArgs e)
        {
            if (ComponentManager.Ribbon == null) return;   // лента ещё не готова — ждём следующий Idle
            AcApp.Idle -= OnIdle;
            foreach (Document d in AcApp.DocumentManager) LoadLisp(d, false);
            SxRibbon.Ensure();
        }

        internal static void LoadLisp(Document doc, bool force)
        {
            if (doc == null) return;
            string lsp = Path.Combine(Root, "CAD", "SX_Tools.lsp").Replace('\\', '/');
            string cmd = "(progn (setq sxt:root \"" + Root.Replace('\\', '/') + "\" sxt:force " + (force ? "T" : "nil")
                       + ") (load \"" + lsp + "\") (princ)) ";
            doc.SendStringToExecute(cmd, false, false, false);
        }

        /// <summary>Выполнить команду в текущем чертеже.
        /// Прерывание (Esc) — только если идёт другая команда и отдельной посылкой:
        /// «^C^C» в одной строке с командой AutoCAD сбрасывает вместе с остатком строки, и команда теряется.</summary>
        internal static void Run(string command)
        {
            Document doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            string active = "";
            try { active = Convert.ToString(AcApp.GetSystemVariable("CMDNAMES")); } catch { }
            if (!string.IsNullOrEmpty(active)) doc.SendStringToExecute("\x1B\x1B", true, false, false);
            doc.SendStringToExecute(command + " ", true, false, true);
        }

        internal static void OpenPath(string path)
        {
            try
            {
                if (!File.Exists(path) && !Directory.Exists(path)) { MessageBox.Show("Не найдено:\n" + path, "SX"); return; }
                Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
            }
            catch (System.Exception ex) { MessageBox.Show(ex.Message, "SX"); }
        }
    }

    public class Commands
    {
        /// <summary>Вернуть и открыть вкладку SX, если её закрыли.</summary>
        [CommandMethod("SXT")]
        public void Show() { SxRibbon.Ensure(); SxRibbon.Activate(); }

        [CommandMethod("SXTRELOAD")]
        public void Reload() { Plugin.LoadLisp(AcApp.DocumentManager.MdiActiveDocument, true); }
    }

    /// <summary>Кнопка ленты: команда AutoCAD или действие.</summary>
    internal class Btn
    {
        public string Title, Command, Tip, Glyph;
        public Action Custom;
        public bool Large;
        public static Btn L(string t, string c, string g, string tip) { return new Btn { Title = t, Command = c, Glyph = g, Tip = tip, Large = true }; }
        public static Btn S(string t, string c, string g, string tip) { return new Btn { Title = t, Command = c, Glyph = g, Tip = tip }; }
        public static Btn A(string t, Action a, string g, string tip) { return new Btn { Title = t, Custom = a, Glyph = g, Tip = tip }; }
    }

    internal class SxCommandHandler : ICommand
    {
        public event EventHandler CanExecuteChanged { add { } remove { } }
        public bool CanExecute(object p) { return true; }
        public void Execute(object p)
        {
            RibbonButton rb = p as RibbonButton;
            Btn b = rb == null ? null : rb.Tag as Btn;
            if (b == null) return;
            if (b.Custom != null) b.Custom(); else Plugin.Run(b.Command);
        }
    }

    internal static class SxRibbon
    {
        const string TabId = "SX_TOOLS_TAB";
        static readonly SxCommandHandler Handler = new SxCommandHandler();
        static bool _dark = true;

        // Значки — Segoe MDL2 Assets (есть в Windows 10/11).
        const string gAdd = "", gRefresh = "", gSelect = "", gPage = "", gBlocks = "",
                     gErase = "", gEdit = "", gCut = "", gDelete = "", gRuler = "",
                     gList = "", gCheck = "", gFile = "", gCrop = "", gPrev = "",
                     gNext = "", gLibrary = "", gSearch = "", gHome = "", gDoc = "",
                     gHelp = "";

        internal static void Ensure()
        {
            RibbonControl ribbon = ComponentManager.Ribbon;
            if (ribbon == null) return;
            foreach (RibbonTab t in ribbon.Tabs) if (t.Id == TabId) return;
            try { _dark = Convert.ToInt32(AcApp.GetSystemVariable("COLORTHEME")) == 0; } catch { }
            ribbon.Tabs.Add(BuildTab());
        }

        internal static void Activate()
        {
            RibbonControl ribbon = ComponentManager.Ribbon;
            if (ribbon == null) return;
            foreach (RibbonTab t in ribbon.Tabs) if (t.Id == TabId) { t.IsVisible = true; t.IsActive = true; }
        }

        static RibbonTab BuildTab()
        {
            string root = Plugin.Root;
            RibbonTab tab = new RibbonTab { Title = "SX", Id = TabId };

            tab.Panels.Add(Panel("Однолинейка",
                new[] {
                    Btn.L("Нарисовать", "SXDRAW", gAdd, "SXDRAW — построить схему: книга, библиотека блоков, лист"),
                    Btn.L("Обновить", "SXUPDATE", gRefresh, "SXUPDATE — перерисовать по книге"),
                    Btn.S("Выбранное", "SXUPDSEL", gSelect, "SXUPDSEL — обновить только выделенные места"),
                    Btn.S("Лист", "SXUPDSHEET", gPage, "SXUPDSHEET — обновить лист целиком"),
                },
                new[] {
                    Btn.S("Блоки", "SXBLOCKS", gBlocks, "SXBLOCKS"),
                    Btn.S("Очистить", "SXCLEAR", gErase, "SXCLEAR — стереть схему"),
                }));

            tab.Panels.Add(Panel("Компоновка",
                new[] {
                    Btn.L("Нарисовать", "CPDRAW", gAdd, "CPDRAW — аппараты и клеммы из блоков"),
                    Btn.L("Обновить", "CPUPDATE", gRefresh, "CPUPDATE"),
                    Btn.S("Править", "CPEDIT", gEdit, "CPEDIT"),
                },
                new[] {
                    Btn.S("Разбить", "CPEXPLODE", gCut, "CPEXPLODE"),
                    Btn.S("Удалить", "CPDELETE", gDelete, "CPDELETE"),
                }));

            tab.Panels.Add(Panel("Длины трасс",
                new[] {
                    Btn.L("Заполнить", "DLN", gRuler, "DLN — длины трасс с планов в книгу"),
                    Btn.L("Обновить", "DLN_UPD", gRefresh, "DLN_UPD — после правок планов"),
                    Btn.S("Сверка", "DLN_SVR", gList, "DLN_SVR — план и таблица по всем листам"),
                    Btn.S("Проверка", "DLN_CHK", gCheck, "DLN_CHK — только проверка"),
                    Btn.S("Книга", "DLN_XLS", gFile, "DLN_XLS — выбрать книгу Excel"),
                },
                new[] {
                    Btn.S("Рамка", "DLN_RAMKA", gCrop, "DLN_RAMKA — выбор рамки → книга"),
                    Btn.S("Пред.", "DLN_PREV", gPrev, "DLN_PREV — к предыдущей пометке"),
                    Btn.S("След.", "DLN_NEXT", gNext, "DLN_NEXT — к следующей пометке"),
                    Btn.S("Сбросить", "DLN_CLR", gErase, "DLN_CLR — убрать пометки"),
                }));

            tab.Panels.Add(Panel("Каталог",
                new[] {
                    Btn.L("Каталог", "SXCAT", gLibrary, "SXCAT — окно каталога блоков"),
                    Btn.S("К блоку", "SXGO", gSearch, "SXGO"),
                    Btn.S("Домой", "SXHOME", gHome, "SXHOME"),
                }, null));

            tab.Panels.Add(Panel("Сервис",
                new[] {
                    Btn.A("Каб. журнал", () => Plugin.OpenPath(Path.Combine(root, "Tools", "CableJournal.exe")), gDoc, "Программа «Кабельный журнал»"),
                    Btn.A("Образец КЖ", () => Plugin.OpenPath(Path.Combine(root, "Tools", "КЖ_шаблон.dwg")), gPage, "КЖ_шаблон.dwg"),
                    Btn.A("Инструкции", () => Plugin.OpenPath(Path.Combine(root, "Docs")), gHelp, "Папка с инструкциями"),
                },
                new[] {
                    Btn.A("Перезагрузить LISP", () => Plugin.LoadLisp(AcApp.DocumentManager.MdiActiveDocument, true), gRefresh, "SXTRELOAD — после правки .lsp"),
                }));

            return tab;
        }

        /// <summary>Панель: крупные кнопки, мелкие столбиками по 3, редкие — в выпадающую часть (стрелка внизу панели).</summary>
        static RibbonPanel Panel(string title, Btn[] main, Btn[] more)
        {
            RibbonPanelSource src = new RibbonPanelSource { Title = title };
            AddButtons(src, main);
            if (more != null && more.Length > 0)
            {
                src.Items.Add(new RibbonPanelBreak());
                AddButtons(src, more);
            }
            return new RibbonPanel { Source = src };
        }

        static void AddButtons(RibbonPanelSource src, Btn[] list)
        {
            RibbonRowPanel col = null;
            int inCol = 0;
            foreach (Btn b in list)
            {
                if (b.Large) { col = null; src.Items.Add(Make(b)); continue; }
                if (col == null || inCol == 3) { col = new RibbonRowPanel(); src.Items.Add(col); inCol = 0; }
                if (inCol > 0) col.Items.Add(new RibbonRowBreak());
                col.Items.Add(Make(b));
                inCol++;
            }
        }

        static RibbonButton Make(Btn b)
        {
            RibbonButton rb = new RibbonButton
            {
                Text = b.Title,
                ShowText = true,
                ShowImage = true,
                Tag = b,
                CommandHandler = Handler,
                Id = "SX_" + (b.Command ?? b.Title),
                ToolTip = new RibbonToolTip { Title = b.Title, Content = b.Tip, Command = b.Command ?? "" },
            };
            if (b.Large)
            {
                rb.Size = RibbonItemSize.Large;
                rb.Orientation = System.Windows.Controls.Orientation.Vertical;
                rb.LargeImage = Glyph(b.Glyph, 32, true);
                rb.Image = Glyph(b.Glyph, 16, true);
            }
            else
            {
                rb.Size = RibbonItemSize.Standard;
                rb.Orientation = System.Windows.Controls.Orientation.Horizontal;
                rb.Image = Glyph(b.Glyph, 16, false);
            }
            return rb;
        }

        /// <summary>Значок из шрифта: акцент для крупных кнопок, нейтральный цвет темы — для мелких.</summary>
        static ImageSource Glyph(string g, int size, bool accent)
        {
            Color c = accent ? Color.FromRgb(0x3D, 0x9B, 0xE9)
                    : (_dark ? Color.FromRgb(0xD8, 0xD8, 0xD8) : Color.FromRgb(0x40, 0x40, 0x40));
            DrawingVisual dv = new DrawingVisual();
            using (DrawingContext dc = dv.RenderOpen())
            {
                FormattedText ft = new FormattedText(g, CultureInfo.InvariantCulture, FlowDirection.LeftToRight,
                    new Typeface("Segoe MDL2 Assets"), size * 0.8, new SolidColorBrush(c), 1.0);
                dc.DrawText(ft, new Point((size - ft.Width) / 2, (size - ft.Height) / 2));
            }
            RenderTargetBitmap bmp = new RenderTargetBitmap(size, size, 96, 96, PixelFormats.Pbgra32);
            bmp.Render(dv);
            bmp.Freeze();
            return bmp;
        }
    }
}
