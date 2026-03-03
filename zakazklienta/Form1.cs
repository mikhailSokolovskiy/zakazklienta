using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;

namespace zakazklienta;

public partial class Form1 : Form
{
    // Колонки (0-based):
    // A=0, B=1, C=2 — мета-данные формата (не данные документа!)
    // D=3 начинается реальный документ
    private const int GROUP_COL = 5; // F — номер заказа
    private const int PROD_START = 65; // BN — начало блока товаров
    private const int COMMENT_COL = 84; // CG — Комментарий

    private string? _excelPath;
    private List<string[]>? _rows;

    // Словарь замены: Артикул -> новая Ссылка номенклатуры
    // Заполняется из дополнительного Excel файла (необязательно)
    private Dictionary<string, string> _refMap = new(StringComparer.Ordinal);

    private readonly Button _btnOpen = new()
        { Text = "📂 Открыть Excel", Width = 170, Height = 36, Left = 12, Top = 12 };

    private readonly Button _btnLoadRef = new()
        { Text = "🔄 Загрузить замены", Width = 170, Height = 36, Left = 192, Top = 12 };

    private readonly Button _btnExport = new()
        { Text = "💾 Сохранить XML", Width = 170, Height = 36, Left = 372, Top = 12, Enabled = false };

    private readonly Label _lblFile = new()
        { Text = "Файл данных: не выбран", Left = 12, Top = 56, Width = 760, Height = 20, AutoSize = false };

    private readonly Label _lblRef = new()
    {
        Text = "Файл замен: не загружен", Left = 12, Top = 78, Width = 760, Height = 20, AutoSize = false,
        ForeColor = System.Drawing.Color.Gray
    };

    private readonly Label _lblStatus = new()
    {
        Left = 12, Top = 100, Width = 760, Height = 20, AutoSize = false, ForeColor = System.Drawing.Color.DarkGreen
    };

    private readonly ListBox _lstOrders = new()
        { Left = 12, Top = 128, Width = 760, Height = 400, Font = new System.Drawing.Font("Consolas", 9f) };

    public Form1()
    {
        Text = "Excel → XML (1С ЗаказКлиента)";
        ClientSize = new System.Drawing.Size(800, 560);
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;
        Controls.AddRange(
            new Control[] { _btnOpen, _btnLoadRef, _btnExport, _lblFile, _lblRef, _lblStatus, _lstOrders });
        _btnOpen.Click += BtnOpen_Click;
        _btnLoadRef.Click += BtnLoadRef_Click;
        _btnExport.Click += BtnExport_Click;
    }

    private void BtnOpen_Click(object? sender, EventArgs e)
    {
        using var dlg = new OpenFileDialog
        {
            Title = "Выберите Excel файл",
            Filter = "Excel файлы (*.xls;*.xlsx)|*.xls;*.xlsx|Все файлы (*.*)|*.*"
        };
        if (dlg.ShowDialog() != DialogResult.OK) return;

        _excelPath = dlg.FileName;
        _lblFile.Text = _excelPath;
        SetStatus("Загрузка...", false);
        _lstOrders.Items.Clear();

        try
        {
            _rows = ReadExcel(_excelPath);
            var groups = GroupByOrder(_rows);

            _lstOrders.Items.Add($"Строк данных: {_rows.Count}   |   Уникальных заказов: {groups.Count}");
            _lstOrders.Items.Add(new string('-', 90));
            foreach (var g in groups)
                _lstOrders.Items.Add($"Заказ № {g.Key}  ->  {g.Value.Count} товар(ов)");

            _btnExport.Enabled = true;
            SetStatus($"✔ Загружено {groups.Count} заказов", false);
        }
        catch (Exception ex)
        {
            SetStatus("Ошибка: " + ex.Message, true);
        }
    }

    // ─── Загрузить файл замен (необязательно) ─────────────────────────────
    // Ожидаемая структура: строка 1 — заголовки, далее данные.
    // Колонка 3 (C, 0-based=2): Уникальный идентификатор (новая Ссылка)
    // Колонка 5 (E, 0-based=4): Артикул (ключ поиска)
    private void BtnLoadRef_Click(object? sender, EventArgs e)
    {
        using var dlg = new OpenFileDialog
        {
            Title = "Выберите Excel файл с заменами ссылок",
            Filter = "Excel файлы (*.xls;*.xlsx)|*.xls;*.xlsx|Все файлы (*.*)|*.*"
        };
        if (dlg.ShowDialog() != DialogResult.OK) return;

        try
        {
            _refMap = ReadRefMap(dlg.FileName);
            _lblRef.Text = $"Файл замен: {dlg.FileName}  ({_refMap.Count} записей)";
            _lblRef.ForeColor = System.Drawing.Color.DarkGreen;
            SetStatus($"✔ Загружено {_refMap.Count} замен артикулов", false);
        }
        catch (Exception ex)
        {
            SetStatus("Ошибка загрузки замен: " + ex.Message, true);
        }
    }

    private void BtnExport_Click(object? sender, EventArgs e)
    {
        if (_rows == null) return;

        using var dlg = new SaveFileDialog
        {
            Title = "Сохранить XML",
            Filter = "XML файлы (*.xml)|*.xml",
            FileName = $"ГОТОВО {Path.GetFileNameWithoutExtension(_excelPath)}.xml",
            DefaultExt = "xml"
        };
        if (dlg.ShowDialog() != DialogResult.OK) return;

        try
        {
            File.WriteAllText(dlg.FileName, BuildXml(_rows, _refMap), new UTF8Encoding(false));
            SetStatus($"✔ Сохранено: {dlg.FileName}", false);
            MessageBox.Show($"Готово!\n{dlg.FileName}", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            SetStatus("Ошибка: " + ex.Message, true);
            MessageBox.Show(ex.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    // ─── Чтение файла замен: возвращает словарь Артикул -> Ссылка ──────────
    private static Dictionary<string, string> ReadRefMap(string path)
    {
        IWorkbook workbook;
        using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            workbook = path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                ? new XSSFWorkbook(fs)
                : new HSSFWorkbook(fs);
        }

        var sheet = workbook.GetSheetAt(1);
        var formatter = new DataFormatter();
        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        int lastRow = sheet.LastRowNum;

        for (int r = 1; r <= lastRow; r++) // строка 0 — заголовок
        {
            var row = sheet.GetRow(r);
            if (row == null) continue;

            // col 2 (C) = Уникальный идентификатор (новая Ссылка)
            // col 4 (E) = Артикул
            string newRef = formatter.FormatCellValue(row.GetCell(2))?.Trim() ?? string.Empty;
            string article = formatter.FormatCellValue(row.GetCell(4))?.Trim() ?? string.Empty;

            if (!string.IsNullOrEmpty(article) && !string.IsNullOrEmpty(newRef))
                map[article] = newRef;
        }

        return map;
    }

    // ─── Чтение Excel: все ячейки как текст через DataFormatter ────────────
    // DataFormatter возвращает именно то, что видно в Excel — без потери
    // точности у длинных чисел (номера счетов, GUID и т.п.)
    private static List<string[]> ReadExcel(string path)
    {
        IWorkbook workbook;
        using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            workbook = path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                ? new XSSFWorkbook(fs)
                : new HSSFWorkbook(fs);
        }

        var sheet = workbook.GetSheetAt(0);
        var formatter = new DataFormatter();
        var result = new List<string[]>();

        int lastRow = sheet.LastRowNum;

        for (int r = 1; r <= lastRow; r++) // строка 0 — заголовок
        {
            var row = sheet.GetRow(r);
            if (row == null) continue;

            int lastCol = row.LastCellNum;
            if (lastCol <= 0) continue;

            var cells = new string[lastCol];
            for (int c = 0; c < lastCol; c++)
            {
                var cell = row.GetCell(c);
                cells[c] = cell == null ? string.Empty : ReadCell(cell, formatter);
            }

            if (cells.Any(v => !string.IsNullOrWhiteSpace(v)))
                result.Add(cells);
        }

        return result;
    }

    // ─── Группировка по колонке F ──────────────────────────────────────────
    private static Dictionary<string, List<string[]>> GroupByOrder(List<string[]> rows)
    {
        var result = new Dictionary<string, List<string[]>>(StringComparer.Ordinal);
        var order = new List<string>();
        foreach (var row in rows)
        {
            string key = Get(row, GROUP_COL);
            if (!result.ContainsKey(key))
            {
                result[key] = new();
                order.Add(key);
            }

            result[key].Add(row);
        }

        var ordered = new Dictionary<string, List<string[]>>(StringComparer.Ordinal);
        foreach (var k in order) ordered[k] = result[k];
        return ordered;
    }

    /// Читает ячейку как строку, для больших целых чисел избегает научной нотации
    private static string ReadCell(ICell cell, DataFormatter formatter)
    {
        if (cell.CellType == CellType.Numeric && !DateUtil.IsCellDateFormatted(cell))
        {
            double d = cell.NumericCellValue;
            if (d == Math.Floor(d) && !double.IsInfinity(d))
                return ((decimal)d).ToString("0"); // decimal не переполняется на таких числах
            return d.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        return formatter.FormatCellValue(cell)?.Trim() ?? string.Empty;
    }

    // ─── Один большой XML со всеми заказами ───────────────────────────────
    private static string BuildXml(List<string[]> rows, Dictionary<string, string> refMap)
    {
        var groups = GroupByOrder(rows);
        var sb = new StringBuilder();
        var settings = new XmlWriterSettings
        {
            Indent = true,
            IndentChars = "  ",
            Encoding = new UTF8Encoding(false),
            OmitXmlDeclaration = false
        };

        using var xw = XmlWriter.Create(sb, settings);
        xw.WriteStartDocument();

        xw.WriteStartElement("Message", "");
        xw.WriteAttributeString("xmlns", "msg", null, "http://www.1c.ru/SSL/Exchange/Message");
        xw.WriteAttributeString("xmlns", "xs", null, "http://www.w3.org/2001/XMLSchema");
        xw.WriteAttributeString("xmlns", "xsi", null, "http://www.w3.org/2001/XMLSchema-instance");

        // Header — берём значения из колонок A, B, C первой строки данных
        var firstRow = rows[0];
        xw.WriteStartElement("msg", "Header", "http://www.1c.ru/SSL/Exchange/Message");
        xw.WriteAttributeString("xmlns", "ns4", null, "http://www.v8.1c.ru/ssl/contactinfo");
        xw.WriteAttributeString("xmlns", "ns3", null, "http://v8.1c.ru/edi/edi_stnd/EnterpriseData/1.0");
        NsE(xw, "msg", "Format", "http://www.1c.ru/SSL/Exchange/Message", Get(firstRow, 0));
        NsE(xw, "msg", "CreationDate", "http://www.1c.ru/SSL/Exchange/Message", Get(firstRow, 1));
        NsE(xw, "msg", "AvailableVersion", "http://www.1c.ru/SSL/Exchange/Message", Get(firstRow, 2));
        xw.WriteEndElement();

        xw.WriteStartElement("Body", "http://v8.1c.ru/edi/edi_stnd/EnterpriseData/1.17");
        foreach (var g in groups)
            WriteOrder(xw, g.Value, refMap);
        xw.WriteEndElement(); // Body

        xw.WriteEndElement(); // Message
        xw.WriteEndDocument();
        xw.Flush();
        return sb.ToString();
    }

    // ─── Один Документ.ЗаказКлиента ───────────────────────────────────────
    // Маппинг колонок (0-based), колонки A=0,B=1,C=2 — мета-данные Header,
    // реальные данные документа начинаются с D=3:
    //
    // D=3   КлючевыеСвойства/Ссылка
    // E=4   КлючевыеСвойства/Дата
    // F=5   КлючевыеСвойства/Номер          ← GROUP_COL
    // G=6   Организация/Ссылка
    // H=7   Организация/Наименование
    // I=8   Организация/НаименованиеСокращенное
    // J=9   Организация/НаименованиеПолное
    // K=10  Организация/ИНН
    // L=11  Организация/КПП
    // M=12  Организация/ЮридическоеФизическоеЛицо
    // N=13  Валюта/Ссылка
    // O=14  Валюта/Код
    // P=15  Валюта/Наименование
    // Q=16  Сумма
    // R=17  Склад/Ссылка
    // S=18  Склад/Наименование
    // T=19  Склад/ТипСклада
    // U=20  Контрагент/Ссылка
    // V=21  Контрагент/Наименование
    // W=22  Контрагент/НаименованиеПолное
    // X=23  Контрагент/ИНН
    // Y=24  Контрагент/КПП
    // Z=25  Контрагент/ЮридическоеФизическоеЛицо
    // AA=26 Договор/Ссылка
    // AB=27 Договор/ВидДоговора
    // AC=28 Договор/Организация/Ссылка
    // AD=29 Договор/Организация/Наименование
    // AE=30 Договор/Организация/НаименованиеСокращенное
    // AF=31 Договор/Организация/НаименованиеПолное
    // AG=32 Договор/Организация/ИНН
    // AH=33 Договор/Организация/КПП
    // AI=34 Договор/Организация/ЮридическоеФизическоеЛицо
    // AJ=35 Договор/Контрагент/Ссылка
    // AK=36 Договор/Контрагент/Наименование
    // AL=37 Договор/Контрагент/НаименованиеПолное
    // AM=38 Договор/Контрагент/ИНН
    // AN=39 Договор/Контрагент/КПП
    // AO=40 Договор/Контрагент/ЮридическоеФизическоеЛицо
    // AP=41 Договор/ВалютаВзаиморасчетов/Ссылка
    // AQ=42 Договор/ВалютаВзаиморасчетов/Код
    // AR=43 Договор/ВалютаВзаиморасчетов/Наименование
    // AS=44 Договор/Наименование
    // AT=45 Договор/Дата
    // AU=46 ДанныеВзаиморасчетов/ВалютаВзаиморасчетов/Ссылка
    // AV=47 ДанныеВзаиморасчетов/ВалютаВзаиморасчетов/Код
    // AW=48 ДанныеВзаиморасчетов/ВалютаВзаиморасчетов/Наименование
    // AX=49 КурсВзаиморасчетов
    // AY=50 КратностьВзаиморасчетов
    // AZ=51 РасчетыВУсловныхЕдиницах
    // BA=52 СуммаВключаетНДС
    // BB=53 НомерСчета
    // BC=54 Банк/Ссылка
    // BD=55 Банк/Наименование
    // BE=56 Банк/БИК
    // BF=57 Банк/КоррСчет
    // BG=58 Владелец/Ссылка
    // BH=59 Владелец/Наименование
    // BI=60 Владелец/НаименованиеСокращенное
    // BJ=61 Владелец/НаименованиеПолное
    // BK=62 Владелец/ИНН
    // BL=63 Владелец/КПП
    // BM=64 Владелец/ЮридическоеФизическоеЛицо
    // BN=65 НомерСтрокиДокумента (auto, ignored)
    // BO=66 Номенклатура/Ссылка         ...далее товарный блок
    private static void WriteOrder(XmlWriter xw, List<string[]> rows, Dictionary<string, string> refMap)
    {
        var f = rows[0];

        xw.WriteStartElement("Документ.ЗаказКлиента", "http://v8.1c.ru/edi/edi_stnd/EnterpriseData/1.17");
        xw.WriteAttributeString("xmlns", "msg", null, "http://www.1c.ru/SSL/Exchange/Message");
        xw.WriteAttributeString("xmlns", "ns3", null, "http://v8.1c.ru/edi/edi_stnd/EnterpriseData/1.0");

        // КлючевыеСвойства
        xw.WriteStartElement("КлючевыеСвойства");
        E(xw, "Ссылка", f, 3);
        E(xw, "Дата", f, 4);
        E(xw, "Номер", f, 5);
        xw.WriteStartElement("Организация");
        E(xw, "Ссылка", f, 6);
        E(xw, "Наименование", f, 7);
        E(xw, "НаименованиеСокращенное", f, 8);
        E(xw, "НаименованиеПолное", f, 9);
        E(xw, "ИНН", f, 10);
        E(xw, "КПП", f, 11);
        E(xw, "ЮридическоеФизическоеЛицо", f, 12);
        xw.WriteEndElement();
        xw.WriteEndElement(); // Организация, КлючевыеСвойства

        // Валюта
        xw.WriteStartElement("Валюта");
        E(xw, "Ссылка", f, 13);
        xw.WriteStartElement("ДанныеКлассификатора");
        E(xw, "Код", f, 14);
        E(xw, "Наименование", f, 15);
        xw.WriteEndElement();
        xw.WriteEndElement();

        EN(xw, "Сумма", f, 16); // Сумма — с округлением

        // Склад
        xw.WriteStartElement("Склад");
        E(xw, "Ссылка", f, 17);
        E(xw, "Наименование", f, 18);
        E(xw, "ТипСклада", f, 19);
        xw.WriteEndElement();

        // Контрагент
        xw.WriteStartElement("Контрагент");
        E(xw, "Ссылка", f, 20);
        E(xw, "Наименование", f, 21);
        E(xw, "НаименованиеПолное", f, 22);
        E(xw, "ИНН", f, 23);
        E(xw, "КПП", f, 24);
        E(xw, "ЮридическоеФизическоеЛицо", f, 25);
        xw.WriteEndElement();

        // ДанныеВзаиморасчетов
        xw.WriteStartElement("ДанныеВзаиморасчетов");

        var contractRef = Get(f, 26);

        if (!string.IsNullOrWhiteSpace(contractRef) && contractRef != "null")
        {
            xw.WriteStartElement("Договор");
            E(xw, "Ссылка", f, 26);
            E(xw, "ВидДоговора", f, 27);
            xw.WriteStartElement("Организация");
            E(xw, "Ссылка", f, 28);
            E(xw, "Наименование", f, 29);
            E(xw, "НаименованиеСокращенное", f, 30);
            E(xw, "НаименованиеПолное", f, 31);
            E(xw, "ИНН", f, 32);
            E(xw, "КПП", f, 33);
            E(xw, "ЮридическоеФизическоеЛицо", f, 34);
            xw.WriteEndElement();
            xw.WriteStartElement("Контрагент");
            E(xw, "Ссылка", f, 35);
            E(xw, "Наименование", f, 36);
            E(xw, "НаименованиеПолное", f, 37);
            E(xw, "ИНН", f, 38);
            E(xw, "КПП", f, 39);
            E(xw, "ЮридическоеФизическоеЛицо", f, 40);
            xw.WriteEndElement();
            xw.WriteStartElement("ВалютаВзаиморасчетов");
            E(xw, "Ссылка", f, 41);
            xw.WriteStartElement("ДанныеКлассификатора");
            E(xw, "Код", f, 42);
            E(xw, "Наименование", f, 43);
            xw.WriteEndElement();
            xw.WriteEndElement();
            E(xw, "Наименование", f, 44);
            xw.WriteElementString("Дата", FormatContractDate(Get(f, 45)));
            xw.WriteEndElement(); // Договор
        }
        
        xw.WriteStartElement("ВалютаВзаиморасчетов");
        E(xw, "Ссылка", f, 46);
        xw.WriteStartElement("ДанныеКлассификатора");
        E(xw, "Код", f, 47);
        E(xw, "Наименование", f, 48);
        xw.WriteEndElement();
        xw.WriteEndElement();
        E(xw, "КурсВзаиморасчетов", f, 49);
        E(xw, "КратностьВзаиморасчетов", f, 50);
        E(xw, "РасчетыВУсловныхЕдиницах", f, 51);
        xw.WriteEndElement(); // ДанныеВзаиморасчетов

        E(xw, "СуммаВключаетНДС", f, 52);

        // БанковскийСчетОрганизации
        xw.WriteStartElement("БанковскийСчетОрганизации");
        E(xw, "НомерСчета", f, 53);
        xw.WriteStartElement("Банк");
        E(xw, "Ссылка", f, 54);
        xw.WriteStartElement("ДанныеКлассификатораБанков");
        E(xw, "Наименование", f, 55);
        E(xw, "БИК", f, 56);
        E(xw, "КоррСчет", f, 57);
        xw.WriteEndElement();
        xw.WriteEndElement();
        xw.WriteStartElement("Владелец");
        xw.WriteStartElement("ОрганизацииСсылка");
        E(xw, "Ссылка", f, 58);
        E(xw, "Наименование", f, 59);
        E(xw, "НаименованиеСокращенное", f, 60);
        E(xw, "НаименованиеПолное", f, 61);
        E(xw, "ИНН", f, 62);
        E(xw, "КПП", f, 63);
        E(xw, "ЮридическоеФизическоеЛицо", f, 64);
        xw.WriteEndElement();
        xw.WriteEndElement();
        xw.WriteEndElement(); // БанковскийСчетОрганизации

        // Товары — все строки группы
        xw.WriteStartElement("Товары");
        for (int i = 0; i < rows.Count; i++)
        {
            var r = rows[i];
            int b = PROD_START; // = 65 (BN)

            // Артикул находится на b+4, используем его для поиска замены Ссылки
            string article = Get(r, b + 4);
            string refFromMap = refMap.TryGetValue(article, out var mapped) ? mapped : Get(r, b + 1);
            // Если артикул найден в словаре — берём новую Ссылку, иначе оригинальную

            xw.WriteStartElement("Строка");
            xw.WriteElementString("НомерСтрокиДокумента", i.ToString()); // нумерация с 0

            xw.WriteStartElement("ДанныеНоменклатуры");
            xw.WriteStartElement("Номенклатура");
            xw.WriteElementString("Ссылка", refFromMap);
            E(xw, "НаименованиеПолное", r, b + 2);
            E(xw, "КодВПрограмме", r, b + 3);
            E(xw, "Артикул", r, b + 4);
            E(xw, "Наименование", r, b + 5);
            xw.WriteEndElement();
            xw.WriteEndElement(); // Номенклатура, ДанныеНоменклатуры

            xw.WriteStartElement("ЕдиницаИзмерения");
            E(xw, "Ссылка", r, b + 6);
            xw.WriteStartElement("ДанныеКлассификатора");
            E(xw, "Код", r, b + 7);
            E(xw, "Наименование", r, b + 8);
            xw.WriteEndElement();
            xw.WriteEndElement();

            E(xw, "Количество", r, b + 9);
            EN(xw, "Сумма", r, b + 10);
            EN(xw, "Цена", r, b + 11);

            xw.WriteStartElement("СтавкаНДС");
            E(xw, "Ставка", r, b + 12);
            E(xw, "РасчетнаяСтавка", r, b + 13);
            E(xw, "НеОблагается", r, b + 14);
            E(xw, "ВидСтавки", r, b + 15);
            xw.WriteStartElement("Страна");
            xw.WriteStartElement("ДанныеКлассификатора");
            E(xw, "Код", r, b + 16);
            E(xw, "Наименование", r, b + 17);
            xw.WriteEndElement();
            xw.WriteEndElement(); // ДанныеКлассификатора, Страна
            xw.WriteEndElement(); // СтавкаНДС

            EN(xw, "СуммаНДС", r, b + 18);
            xw.WriteEndElement(); // Строка
        }

        xw.WriteEndElement(); // Товары

        xw.WriteStartElement("ОбщиеСвойстваОбъектовФормата");
        xw.WriteElementString("Комментарий", Get(f, COMMENT_COL));
        xw.WriteEndElement();

        xw.WriteEndElement(); // Документ.ЗаказКлиента
    }
    
    

    // ─── Хелперы ──────────────────────────────────────────────────────────

    
    private static string FormatContractDate(string val)
    {
        if (string.IsNullOrWhiteSpace(val))
            return string.Empty;

        if (val == "null.null.null null:null")
            return string.Empty;

        // ожидаем формат: 16.04.2025 10:42
        var parts = val.Split(' ');
        if (parts.Length != 2)
            return val;

        var dateParts = parts[0].Split('.');
        if (dateParts.Length != 3)
            return val;

        var day = dateParts[0];
        var month = dateParts[1];
        var year = dateParts[2];
        var time = parts[1];

        return $"{year}-{month}-{day}+{time}";
    }
    
    /// Записать элемент — значение как есть (текст, GUID, даты и т.п.)
    private static void E(XmlWriter xw, string name, string[] row, int idx)
        => xw.WriteElementString(name, Get(row, idx));

    /// Записать элемент — числовое значение с округлением до 2 знаков
    private static void EN(XmlWriter xw, string name, string[] row, int idx)
        => xw.WriteElementString(name, RoundNum(Get(row, idx)));

    private static void NsE(XmlWriter xw, string prefix, string local, string ns, string value)
    {
        xw.WriteStartElement(prefix, local, ns);
        xw.WriteString(value);
        xw.WriteEndElement();
    }

    private static string Get(string[] row, int idx)
        => (idx >= 0 && idx < row.Length) ? (row[idx] ?? string.Empty) : string.Empty;

    /// Округляет до 2 десятичных знаков если их больше 2; иначе не трогает
    private static string RoundNum(string val)
    {
        if (string.IsNullOrWhiteSpace(val)) return val;
        // Нормализуем разделитель
        string normalized = val.Replace(',', '.');
        if (!decimal.TryParse(normalized, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out decimal d))
            return val; // не число — оставляем как есть

        int dotPos = normalized.IndexOf('.');
        int decimals = dotPos < 0 ? 0 : normalized.Length - dotPos - 1;

        if (decimals > 2)
            d = Math.Round(d, 2, MidpointRounding.AwayFromZero);

        // Форматируем: убираем лишние нули но оставляем точность
        return d.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
    }

    private void SetStatus(string text, bool isError)
    {
        _lblStatus.Text = text;
        _lblStatus.ForeColor = isError ? System.Drawing.Color.Red : System.Drawing.Color.DarkGreen;
    }
}