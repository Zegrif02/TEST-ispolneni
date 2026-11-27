using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.IO;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace EmbodimentApiProbe;

public partial class MainWindow : Window
{
    private readonly ObservableCollection<ExecutionRow> _rows = new();
    private readonly StringBuilder _log = new();
    private const string KompasProgId = "KOMPAS.Application.7";
    private const int BaseMarkingType = 1; // ksVMBaseMarking
    private const string WriteConfirmationText = "ИЗМЕНИТЬ";
    private static readonly Guid Api7TypeLibGuid = new("69AC2981-37C0-4379-84FD-5DD2F3C0A520");
    private string _lastReadDocumentPath = string.Empty;
    private string _writeAttemptedDocumentPath = string.Empty;
    private bool _hasSuccessfulRead;
    private bool _writeAttempted;
    private bool _pendingPersistenceCheck;
    private string _pendingPersistencePath = string.Empty;
    private List<ExecutionSnapshot> _expectedPersistedRows = new();

    public MainWindow()
    {
        InitializeComponent();
        ResultsGrid.ItemsSource = _rows;
        Log("Диагностическая программа 0.2.2 запущена.");
        Log("Режим чтения доступен всегда. Тест записи и одного перестроения отдельно защищён подтверждениями и требует тестовую копию.");
        Log("После SetEmbodimentMarking выполняется ровно один IKompasDocument3D.RebuildDocument(). Save программой не вызывается.");
        Log("Проверка после перестроения не привязана к индексам: учитывается всё семейство с общей базой и сохранность номеров исполнений.");
    }

    private void ReadButton_Click(object sender, RoutedEventArgs e)
    {
        ReadButton.IsEnabled = false;
        StatusText.Text = "Выполняется...";
        _rows.Clear();
        _lastReadDocumentPath = string.Empty;
        _hasSuccessfulRead = false;
        _pendingPersistenceCheck = false;
        _pendingPersistencePath = string.Empty;
        _expectedPersistedRows.Clear();
        PersistenceCheckBox.IsChecked = false;
        RefreshPersistenceButtonState();
        NewBaseMarkingBox.Clear();
        ConfirmationPhraseBox.Clear();
        TestCopyCheckBox.IsChecked = false;
        RefreshWriteButtonState();
        Log("============================================================");
        Log("Начало диагностики активного документа.");

        try
        {
            bool connected = WithActiveManager((document, doc3d, manager, documentPath) =>
            {
                if (!PathsEqual(documentPath, _writeAttemptedDocumentPath))
                    _writeAttempted = false;
                _lastReadDocumentPath = documentPath;
                Log("IEmbodimentsManager получен. Начинаем чтение.");
                ReadManager(manager);
                _hasSuccessfulRead = _rows.Count > 0;
            });

            StatusText.Text = connected ? $"Готово. Строк: {_rows.Count}" : "Не удалось подключиться — см. журнал";
        }
        catch (Exception ex)
        {
            LogError("Необработанная ошибка чтения", ex);
            StatusText.Text = "Ошибка — подробности в журнале";
        }
        finally
        {
            ReadButton.IsEnabled = true;
            RefreshWriteButtonState();
        }
    }

    /// <summary>
    /// Подключается только к уже запущенному КОМПАС, получает активный документ и IEmbodimentsManager.
    /// Все COM-вызовы выполняются в UI STA-потоке.
    /// </summary>
    private bool WithActiveManager(Action<object, object?, object, string> action)
    {
        object? app = null;
        object? document = null;
        object? doc3d = null;
        object? manager = null;
        IntPtr docPtr = IntPtr.Zero;
        IntPtr doc3dPtr = IntPtr.Zero;
        IntPtr managerPtr = IntPtr.Zero;

        try
        {
            Log("Поиск уже запущенного КОМПАС-3D через ROT...");
            var appType = Type.GetTypeFromProgID(KompasProgId, false);
            if (appType is null)
            {
                Log("ProgID KOMPAS.Application.7 не зарегистрирован.");
                return false;
            }

            app = GetActiveComObject(appType.GUID);
            if (app is null)
            {
                Log("Активный экземпляр КОМПАС-3D не найден. Новый КОМПАС не запускается.");
                return false;
            }

            Log("Подключение к существующему экземпляру успешно.");
            document = ((dynamic)app).ActiveDocument;
            if (document is null)
            {
                Log("В КОМПАС нет активного документа.");
                return false;
            }

            LogDocumentInfo(document);
            string documentPath = ReadDocumentPath(document);
            Log("Загрузка зарегистрированной библиотеки типов API7...");
            var interfaces = TypeLibInterfaceCatalog.FindInterfaces(Api7TypeLibGuid, "IEmbodimentsManager", "IKompasDocument3D");
            Log(interfaces.Report);

            if (!interfaces.InterfaceIds.TryGetValue("IKompasDocument3D", out var document3dIid))
                throw new InvalidOperationException("Не найден IID интерфейса IKompasDocument3D в TypeLib API7.");
            if (!interfaces.InterfaceIds.TryGetValue("IEmbodimentsManager", out var managerIid))
                throw new InvalidOperationException("Не найден IID интерфейса IEmbodimentsManager в TypeLib API7.");

            docPtr = Marshal.GetIUnknownForObject(document);
            int doc3dHr = Marshal.QueryInterface(docPtr, in document3dIid, out doc3dPtr);
            if (doc3dHr >= 0 && doc3dPtr != IntPtr.Zero)
            {
                doc3d = Marshal.GetObjectForIUnknown(doc3dPtr);
                Log("Интерфейс IKompasDocument3D получен через QueryInterface.");
                int managerHr = Marshal.QueryInterface(doc3dPtr, in managerIid, out managerPtr);
                if (managerHr < 0 || managerPtr == IntPtr.Zero)
                {
                    Log("QueryInterface IEmbodimentsManager от IKompasDocument3D не выполнен. Пробуем ActiveDocument...");
                    Marshal.QueryInterface(docPtr, in managerIid, out managerPtr);
                }
            }
            else
            {
                Log("QueryInterface IKompasDocument3D не выполнен. Пробуем IEmbodimentsManager напрямую от ActiveDocument...");
                Marshal.QueryInterface(docPtr, in managerIid, out managerPtr);
            }

            if (managerPtr == IntPtr.Zero)
                throw new InvalidOperationException("Не удалось получить IEmbodimentsManager от активного документа.");

            manager = Marshal.GetObjectForIUnknown(managerPtr);
            action(document, doc3d, manager, documentPath);
            return true;
        }
        finally
        {
            ReleaseCom(manager);
            ReleaseCom(doc3d);
            ReleaseCom(document);
            ReleaseCom(app);
            if (managerPtr != IntPtr.Zero) Marshal.Release(managerPtr);
            if (doc3dPtr != IntPtr.Zero) Marshal.Release(doc3dPtr);
            if (docPtr != IntPtr.Zero) Marshal.Release(docPtr);
        }
    }

    private string ReadDocumentPath(object document)
    {
        try { return Convert.ToString(((dynamic)document).PathName) ?? string.Empty; }
        catch { return string.Empty; }
    }

    private void LogDocumentInfo(object document)
    {
        Log("--- Активный документ ---");
        TryLogProperty(document, "Name");
        TryLogProperty(document, "PathName");
        TryLogProperty(document, "DocumentType");
        TryLogProperty(document, "Type");
    }

    private void ReadManager(object manager)
    {
        int? count = TryReadCount(manager);
        Log($"Количество исполнений: {(count.HasValue ? count.Value.ToString(CultureInfo.InvariantCulture) : "не удалось получить")}");

        // Два необязательных метода GetEmbodimentsTree и GetCurrentEmbodimentIndex
        // намеренно не вызываются: их COM-сигнатуры/параметры не нужны для проверки записи.
        TryInvokeAndLog(manager, "GetCurrentEmbodimentMarking", "Обозначение текущего исполнения (полное)", new object[] { -1, false });

        if (!count.HasValue || count.Value <= 0)
        {
            Log("Количество исполнений не получено или равно нулю. Перебор индексов отменён.");
            return;
        }

        Log("--- Чтение обозначений по индексам (0-based) ---");
        for (int index = 0; index < count.Value; index++)
            ReadOneIndex(manager, index, count.Value);

        Log("--- Дополнительная проверка граничного индекса count (только чтение) ---");
        ReadOneIndex(manager, count.Value, count.Value);
    }

    private void ReadOneIndex(object manager, int index, int count)
    {
        string baseMark = TryGetMarking(manager, index, 1, "Базовая часть");
        string number = TryGetMarking(manager, index, 2, "Номер исполнения");
        string full = TryGetMarking(manager, index, -1, "Полное обозначение");
        string variant = TryGetVariant(manager, index);
        Log($"Индекс {index}: Base='{baseMark}' | Number='{number}' | Full='{full}' | IsVariant='{variant}'");

        // Граничный индекс count оставляем в журнале, но не добавляем в рабочую таблицу.
        if (index < count &&
            baseMark != "<ошибка>" && number != "<ошибка>" && full != "<ошибка>" &&
            (!string.IsNullOrWhiteSpace(baseMark) || !string.IsNullOrWhiteSpace(number) || !string.IsNullOrWhiteSpace(full)))
        {
            _rows.Add(new ExecutionRow
            {
                Index = index,
                BaseMarking = baseMark,
                ExecutionNumber = number,
                FullMarking = full,
                IsVariant = variant
            });
        }
    }

    private int? TryReadCount(object manager)
    {
        foreach (string property in new[] { "EmbodimentCount", "GetEmbodimentCount" })
        {
            try
            {
                object? value = property.StartsWith("Get", StringComparison.Ordinal)
                    ? ((dynamic)manager).GetEmbodimentCount()
                    : ((dynamic)manager).EmbodimentCount;
                if (value is not null && int.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), out int count))
                {
                    Log($"Получено через {property}: {count}");
                    return count;
                }
            }
            catch (Exception ex) { LogError($"Чтение {property}", ex); }
        }
        return null;
    }

    private string TryGetMarking(object manager, int index, int type, string label)
    {
        try
        {
            dynamic d = manager;
            object? result = d.GetEmbodimentMarking(index, type, false);
            return Convert.ToString(result, CultureInfo.InvariantCulture) ?? string.Empty;
        }
        catch (Exception ex)
        {
            LogError($"GetEmbodimentMarking(index={index}, type={type}) [{label}]", ex);
            return "<ошибка>";
        }
    }

    private string TryGetVariant(object manager, int index)
    {
        try
        {
            object? result = ((dynamic)manager).IsVariant(index);
            return Convert.ToString(result, CultureInfo.InvariantCulture) ?? string.Empty;
        }
        catch (Exception ex)
        {
            LogError($"IsVariant(index={index})", ex);
            return "<ошибка>";
        }
    }

    private void TryInvokeAndLog(object manager, string methodName, string label, object[] args)
    {
        try
        {
            dynamic d = manager;
            object? result = methodName switch
            {
                "GetEmbodimentsTree" => d.GetEmbodimentsTree(),
                "GetCurrentEmbodimentIndex" => d.GetCurrentEmbodimentIndex(),
                "GetCurrentEmbodimentMarking" => d.GetCurrentEmbodimentMarking((int)args[0], (bool)args[1]),
                _ => null
            };
            Log($"{label}: {FormatComResult(result)}");
        }
        catch (Exception ex) { LogError(label, ex); }
    }

    private static string FormatComResult(object? value)
    {
        if (value is null) return "<null>";
        if (value is Array array)
        {
            var parts = new List<string>();
            foreach (object? item in array) parts.Add(Convert.ToString(item, CultureInfo.InvariantCulture) ?? "<null>");
            return $"Array[{array.Length}]: " + string.Join(" | ", parts);
        }
        return Convert.ToString(value, CultureInfo.InvariantCulture) ?? value.ToString() ?? "<unknown>";
    }

    private void TryLogProperty(object target, string propertyName)
    {
        try
        {
            object? value = propertyName switch
            {
                "Name" => ((dynamic)target).Name,
                "PathName" => ((dynamic)target).PathName,
                "DocumentType" => ((dynamic)target).DocumentType,
                "Type" => ((dynamic)target).Type,
                _ => null
            };
            Log($"{propertyName}: {Convert.ToString(value, CultureInfo.InvariantCulture) ?? "<null>"}");
        }
        catch (Exception ex) { LogError($"Свойство {propertyName}", ex); }
    }

    private void ResultsGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (ResultsGrid.SelectedItem is ExecutionRow row)
        {
            SelectedRowText.Text = $"Выбрано: индекс {row.Index}; база '{row.BaseMarking}'; исполнение '{row.ExecutionNumber}'; полное '{row.FullMarking}'; вариант геометрии: {row.IsVariant}.";
            if (!string.IsNullOrWhiteSpace(row.ExecutionNumber))
                NewBaseMarkingBox.Text = row.BaseMarking;
        }
        else
        {
            SelectedRowText.Text = "Выбери строку исполнения в таблице. Для теста нужен непустой номер исполнения, например -02.";
        }
        RefreshWriteButtonState();
    }

    private void WriteConfirmation_TextChanged(object sender, TextChangedEventArgs e) => RefreshWriteButtonState();

    private void WriteConfirmation_CheckChanged(object sender, RoutedEventArgs e) => RefreshWriteButtonState();

    private void RefreshWriteButtonState()
    {
        if (WriteTestButton is null) return;
        WriteTestButton.IsEnabled = _hasSuccessfulRead && !_writeAttempted &&
                                    ResultsGrid.SelectedItem is ExecutionRow row &&
                                    !string.IsNullOrWhiteSpace(row.ExecutionNumber) &&
                                    string.Equals(row.IsVariant, "False", StringComparison.OrdinalIgnoreCase) &&
                                    !string.IsNullOrWhiteSpace(NewBaseMarkingBox.Text) &&
                                    TestCopyCheckBox.IsChecked == true &&
                                    string.Equals(ConfirmationPhraseBox.Text.Trim(), WriteConfirmationText, StringComparison.Ordinal);
    }

    private void WriteTestButton_Click(object sender, RoutedEventArgs e)
    {
        if (ResultsGrid.SelectedItem is not ExecutionRow selected)
        {
            MessageBox.Show("Сначала выбери строку в таблице.", "Тест записи", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        string newBaseMarking = NewBaseMarkingBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(newBaseMarking) || newBaseMarking.Contains('\r') || newBaseMarking.Contains('\n'))
        {
            MessageBox.Show("Введи одно непустое базовое обозначение.", "Тест записи", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        if (string.IsNullOrWhiteSpace(selected.ExecutionNumber))
        {
            MessageBox.Show("Выбери исполнение с непустым номером, например -02. Это нужно, чтобы проверить сохранность номера исполнения.", "Тест записи", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        if (string.Equals(selected.IsVariant, "True", StringComparison.OrdinalIgnoreCase))
        {
            MessageBox.Show("Для этого теста варианты геометрии не поддерживаются. Выбери обычное исполнение.", "Тест записи", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        if (string.Equals(newBaseMarking, selected.BaseMarking, StringComparison.Ordinal))
        {
            MessageBox.Show("Новое базовое обозначение совпадает с текущим. Введи другое значение.", "Тест записи", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        if (string.IsNullOrWhiteSpace(_lastReadDocumentPath))
        {
            MessageBox.Show("Сначала считай исполнения из сохранённого тестового документа.", "Тест записи", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        string oldBaseMarking = selected.BaseMarking;
        var beforeRows = _rows.Select(ToSnapshot).ToList();
        var familyBefore = beforeRows.Where(row => string.Equals(row.BaseMarking, oldBaseMarking, StringComparison.Ordinal)).ToList();
        if (familyBefore.Count == 0 || !familyBefore.Any(row => string.Equals(row.ExecutionNumber, selected.ExecutionNumber, StringComparison.Ordinal)))
        {
            MessageBox.Show("Не удалось определить семейство исполнений с общей базой. Тест записи отменён.", "Тест записи", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        if (beforeRows.Any(row => !string.Equals(row.BaseMarking, oldBaseMarking, StringComparison.Ordinal) &&
                                  string.Equals(row.BaseMarking, newBaseMarking, StringComparison.Ordinal)))
        {
            MessageBox.Show("В документе уже есть другое семейство с таким базовым обозначением. Для безопасной проверки выбери уникальную новую базу.", "Тест записи", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // Наблюдение по тесту 0.2.1: после перестроения КОМПАС может переместить строки,
        // а запись базовой части отражается на всех строках семейства с общей базой.
        // Поэтому строим ожидаемое состояние всего семейства, сохраняя каждый номер исполнения.
        var expectedAfterRows = beforeRows.Select(row =>
            string.Equals(row.BaseMarking, oldBaseMarking, StringComparison.Ordinal)
                ? row with { BaseMarking = newBaseMarking, FullMarking = newBaseMarking + row.ExecutionNumber }
                : row).ToList();

        string expectedSelectedFull = newBaseMarking + selected.ExecutionNumber;
        string familyPreview = string.Join(Environment.NewLine, familyBefore.Select(row =>
            $"  {row.FullMarking}  →  {newBaseMarking}{row.ExecutionNumber}"));
        string prompt = "СЕЙЧАС БУДЕТ ИЗМЕНЕНО ОБОЗНАЧЕНИЕ В ОТКРЫТОМ ДОКУМЕНТЕ КОМПАС.\n\n" +
                        $"Файл: {_lastReadDocumentPath}\n" +
                        $"Выбранная строка: индекс {selected.Index}\n" +
                        $"Семейство с общей базой '{oldBaseMarking}': {familyBefore.Count} строк(и)\n" +
                        familyPreview + "\n\n" +
                        $"Новое базовое обозначение семейства: {newBaseMarking}\n" +
                        "Номера исполнений всех строк семейства должны сохраниться без изменений.\n" +
                        $"Для выбранной строки ожидается: {expectedSelectedFull}\n\n" +
                        "Программа вызовет SetEmbodimentMarking(index, ksVMBaseMarking, newBase), один раз выполнит RebuildDocument и проверит весь набор исполнений без привязки к индексам. Save программой не вызывается.\n\n" +
                        "Продолжать только на отдельной одноразовой копии! После успешной проверки сохранение и повторное открытие выполняются вручную в КОМПАС.";

        if (MessageBox.Show(prompt, "Последнее подтверждение изменения", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No) != MessageBoxResult.Yes)
        {
            Log("Пользователь отменил тест записи на последнем подтверждении.");
            return;
        }

        WriteTestButton.IsEnabled = false;
        StatusText.Text = "Выполняется контролируемая запись...";
        Log("============================================================");
        Log("НАЧАЛО ТЕСТА ЗАПИСИ СЕМЕЙСТВА И ОДНОГО ПЕРЕСТРОЕНИЯ. Save не вызывается.");
        Log($"Выбрана строка: индекс {selected.Index}; Base='{selected.BaseMarking}'; Number='{selected.ExecutionNumber}'; Full='{selected.FullMarking}'.");
        Log($"В семейство с базой '{oldBaseMarking}' входит строк: {familyBefore.Count}. После перестроения индексы могут измениться.");
        foreach (var familyRow in familyBefore)
            Log($"Ожидаемое изменение семейства: '{familyRow.FullMarking}' → '{newBaseMarking}{familyRow.ExecutionNumber}' (номер '{familyRow.ExecutionNumber}' сохраняется).");

        bool rebuildSucceeded = false;
        Exception? rebuildException = null;
        bool writeAccepted = false;
        int selectedIndex = selected.Index;
        string selectedNumberBefore = selected.ExecutionNumber;

        try
        {
            bool writeConnected = WithActiveManager((document, doc3d, manager, activePath) =>
            {
                if (!PathsEqual(activePath, _lastReadDocumentPath))
                    throw new InvalidOperationException("Активный документ КОМПАС изменился после чтения. Повторно считай исполнения и проверь выбранную копию.");

                if (doc3d is null)
                    throw new InvalidOperationException("Не получен IKompasDocument3D. Запись отменена, поскольку после неё нельзя безопасно выполнить перестроение документа.");

                int? count = TryReadCount(manager);
                if (!count.HasValue || selectedIndex < 0 || selectedIndex >= count.Value)
                    throw new InvalidOperationException($"Индекс {selectedIndex} больше не входит в диапазон исполнений текущего документа.");

                string beforeBase = TryGetMarking(manager, selectedIndex, 1, "База перед записью");
                string beforeNumber = TryGetMarking(manager, selectedIndex, 2, "Номер перед записью");
                string beforeFull = TryGetMarking(manager, selectedIndex, -1, "Полное обозначение перед записью");
                Log($"Проверка перед записью: Base='{beforeBase}' | Number='{beforeNumber}' | Full='{beforeFull}'");
                if (beforeBase != selected.BaseMarking || beforeNumber != selected.ExecutionNumber || beforeFull != selected.FullMarking)
                    throw new InvalidOperationException("Данные выбранного исполнения изменились после последнего чтения. Запись отменена — нажми считывание ещё раз.");

                // Также убеждаемся, что список до записи совпадает с таблицей, по содержимому, а не по порядку.
                var latestRows = ReadSnapshotsWithoutGrid(manager, count.Value);
                if (!AreMultisetsEqual(beforeRows, latestRows))
                    throw new InvalidOperationException("Состав исполнений изменился после последнего чтения. Запись отменена — обнови таблицу и повтори тест.");

                Log($"Вызов SetEmbodimentMarking(index={selectedIndex}, MarkingType={BaseMarkingType} / ksVMBaseMarking, Marking='{newBaseMarking}')...");
                _writeAttempted = true;
                _writeAttemptedDocumentPath = activePath;
                RefreshWriteButtonState();
                object? apiResult = ((dynamic)manager).SetEmbodimentMarking(selectedIndex, BaseMarkingType, newBaseMarking);
                writeAccepted = apiResult is not null && Convert.ToBoolean(apiResult, CultureInfo.InvariantCulture);
                Log($"Результат SetEmbodimentMarking: {FormatComResult(apiResult)}");
                if (!writeAccepted)
                    throw new InvalidOperationException("SetEmbodimentMarking вернул FALSE. Перестроение не выполняется. Закрой тестовую копию без сохранения.");

                string immediateBase = TryGetMarking(manager, selectedIndex, 1, "База сразу после записи, до перестроения");
                string immediateNumber = TryGetMarking(manager, selectedIndex, 2, "Номер сразу после записи, до перестроения");
                string immediateFull = TryGetMarking(manager, selectedIndex, -1, "Полное обозначение сразу после записи, до перестроения");
                Log($"Снимок до перестроения (может быть неактуальным): Base='{immediateBase}' | Number='{immediateNumber}' | Full='{immediateFull}'");

                Log("Вызов IKompasDocument3D.RebuildDocument() — ровно один раз. Save не вызывается.");
                try
                {
                    object? rebuildResult = ((dynamic)doc3d).RebuildDocument();
                    rebuildSucceeded = rebuildResult is not null && Convert.ToBoolean(rebuildResult, CultureInfo.InvariantCulture);
                    Log($"Результат RebuildDocument: {FormatComResult(rebuildResult)}");
                    if (!rebuildSucceeded)
                        Log("ПРЕДУПРЕЖДЕНИЕ: RebuildDocument вернул FALSE. Выполняем повторное чтение только для диагностики.");
                }
                catch (Exception ex)
                {
                    rebuildException = ex;
                    LogError("Вызов IKompasDocument3D.RebuildDocument", ex);
                }
            });

            if (!writeConnected)
                throw new InvalidOperationException("Не удалось подключиться к активному КОМПАС для записи. Перестроение и повторная проверка не выполнены.");

            Log("--- Повторное подключение к активному документу после RebuildDocument ---");
            _rows.Clear();
            bool verifyConnected = WithActiveManager((document, doc3d, manager, activePath) =>
            {
                if (!PathsEqual(activePath, _lastReadDocumentPath))
                    throw new InvalidOperationException("После перестроения активным стал другой документ. Проверка остановлена.");
                Log("Свежий IEmbodimentsManager получен. Повторно считываем ВСЕ исполнения.");
                ReadManager(manager);
            });

            if (!verifyConnected)
                throw new InvalidOperationException("Не удалось повторно подключиться к документу после перестроения. Не сохраняй тестовую копию.");

            var actualAfterRows = _rows.Select(ToSnapshot).ToList();
            var actualFamilyRows = actualAfterRows.Where(row => string.Equals(row.BaseMarking, newBaseMarking, StringComparison.Ordinal)).ToList();
            var expectedFamilyRows = expectedAfterRows.Where(row => string.Equals(row.BaseMarking, newBaseMarking, StringComparison.Ordinal)).ToList();
            var expectedSelected = new ExecutionSnapshot(newBaseMarking, selectedNumberBefore, expectedSelectedFull, selected.IsVariant);

            bool targetFamilyFound = actualFamilyRows.Any(row => row.Equals(expectedSelected));
            bool baseMatches = targetFamilyFound;
            bool numberPreserved = AreStringMultisetsEqual(
                familyBefore.Select(row => row.ExecutionNumber),
                actualFamilyRows.Select(row => row.ExecutionNumber));
            bool fullMatches = actualFamilyRows.Any(row => row.Equals(expectedSelected));
            bool familyCountMatches = actualFamilyRows.Count == expectedFamilyRows.Count;
            bool rowCountMatches = actualAfterRows.Count == expectedAfterRows.Count;
            bool allRowsMatchExpected = AreMultisetsEqual(expectedAfterRows, actualAfterRows);

            Log("--- Результат сопоставления семейств без использования индексов ---");
            Log($"Старое семейство '{oldBaseMarking}': {familyBefore.Count} строк; новое семейство '{newBaseMarking}': {actualFamilyRows.Count} строк.");
            Log($"Ожидаемые номера семейства: [{string.Join(", ", expectedFamilyRows.Select(row => row.ExecutionNumber))}]");
            Log($"Фактические номера семейства: [{string.Join(", ", actualFamilyRows.Select(row => row.ExecutionNumber))}]");
            foreach (var familyRow in actualFamilyRows)
            {
                var actualIndex = _rows.FirstOrDefault(row => ToSnapshot(row).Equals(familyRow))?.Index;
                Log($"После перестроения: индекс {actualIndex?.ToString(CultureInfo.InvariantCulture) ?? "?"}; Base='{familyRow.BaseMarking}' | Number='{familyRow.ExecutionNumber}' | Full='{familyRow.FullMarking}'");
            }

            string multisetDiff = DescribeMultisetDifference(expectedAfterRows, actualAfterRows);
            Log("--- Итоговая проверка после одного RebuildDocument и повторного подключения ---");
            Log($"Проверки: TargetFamilyFound={targetFamilyFound}; SelectedBaseMatches={baseMatches}; ExecutionNumbersPreserved={numberPreserved}; SelectedFullMatches={fullMatches}; FamilyCountMatches={familyCountMatches}; RowCountMatches={rowCountMatches}; RebuildSucceeded={rebuildSucceeded}; EntireSetMatchesIgnoringIndex={allRowsMatchExpected}");

            bool verificationPassed = baseMatches && numberPreserved && fullMatches && familyCountMatches && rowCountMatches &&
                                     allRowsMatchExpected && rebuildSucceeded && rebuildException is null;
            if (verificationPassed)
            {
                _pendingPersistenceCheck = true;
                _pendingPersistencePath = _lastReadDocumentPath;
                _expectedPersistedRows = expectedAfterRows;
                PersistenceCheckBox.IsChecked = false;
                RefreshPersistenceButtonState();

                Log("УСПЕХ: одно RebuildDocument применило новое базовое обозначение ко всему семейству. Все номера исполнений и остальные семейства сохранены; индексы не использовались для сравнения.");
                Log("Следующий этап: вручную сохрани тестовую копию в КОМПАС, закрой документ, открой тот же файл заново и выполни проверку сохранённого состояния. Программа Save не вызывала.");
                StatusText.Text = "Перестроение успешно — ожидается проверка сохранения";
                MessageBox.Show("Тест перестроения пройден!\n\nКОМПАС изменил базовое обозначение семейства, сохранил номера исполнений, а сравнение всего списка подтвердило, что остальные строки не изменились. Порядок/индексы после перестроения не учитывались.\n\nСледующий этап: вручную СОХРАНИ тестовую копию в КОМПАС, закрой этот документ, затем заново открой тот же файл. Программу EmbodimentApiProbe оставь открытой. После повторного открытия отметь флажок в блоке проверки сохранения и нажми кнопку проверки.\n\nПрограмма сама ничего не сохраняет.",
                    "Перестроение подтверждено", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                string details = $"TargetFamilyFound={targetFamilyFound}; ExecutionNumbersPreserved={numberPreserved}; SelectedFullMatches={fullMatches}; FamilyCountMatches={familyCountMatches}; RowCountMatches={rowCountMatches}; RebuildSucceeded={rebuildSucceeded}; EntireSetMatchesIgnoringIndex={allRowsMatchExpected}.\n" + multisetDiff;
                if (rebuildException is not null) details += $"\nОшибка RebuildDocument: {rebuildException.Message}";
                LogError("ПРОВЕРКА ПОСЛЕ ПЕРЕСТРОЕНИЯ НЕ ПРОЙДЕНА", new InvalidOperationException(details + " Не сохраняй тестовую копию; закрой её без сохранения."));
                StatusText.Text = "Проверка не пройдена — НЕ СОХРАНЯТЬ";
                MessageBox.Show("Проверка после перестроения не пройдена полностью. НЕ СОХРАНЯЙ документ; закрой тестовую копию без сохранения.\n\n" + details,
                    "Тест перестроения", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        catch (Exception ex)
        {
            LogError("Тест записи/перестроения прерван", ex);
            StatusText.Text = "Тест прерван — не сохраняй тестовую копию";
            MessageBox.Show("Тест не завершился штатно. Не сохраняй тестовую копию. Закрой её без сохранения и отправь мне лог.\n\n" + ex.Message,
                "Тест записи и перестроения", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            RefreshWriteButtonState();
        }
    }

    private void PersistenceCheckBox_Changed(object sender, RoutedEventArgs e) => RefreshPersistenceButtonState();

    private void RefreshPersistenceButtonState()
    {
        if (VerifyPersistenceButton is null || PersistenceCheckBox is null) return;
        VerifyPersistenceButton.IsEnabled = _pendingPersistenceCheck && PersistenceCheckBox.IsChecked == true;
    }

    private void VerifyPersistenceButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_pendingPersistenceCheck || string.IsNullOrWhiteSpace(_pendingPersistencePath) || _expectedPersistedRows.Count == 0)
        {
            MessageBox.Show("Нет ожидающего результата теста для проверки сохранения. Сначала успешно выполни тест записи и перестроения.",
                "Проверка сохранения", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        if (PersistenceCheckBox.IsChecked != true)
        {
            MessageBox.Show("Подтверди, что ты вручную сохранил тестовую копию в КОМПАС, закрыл документ и заново открыл тот же файл.",
                "Проверка сохранения", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        VerifyPersistenceButton.IsEnabled = false;
        StatusText.Text = "Проверка повторно открытого файла...";
        Log("============================================================");
        Log("НАЧАЛО ПРОВЕРКИ ПОСЛЕ РУЧНОГО СОХРАНЕНИЯ, ЗАКРЫТИЯ И ПОВТОРНОГО ОТКРЫТИЯ.");
        Log("Программа не вызывает Save, Close или Open. Проверяется только активный документ.");
        _rows.Clear();

        try
        {
            bool connected = WithActiveManager((document, doc3d, manager, activePath) =>
            {
                if (!PathsEqual(activePath, _pendingPersistencePath))
                    throw new InvalidOperationException($"Открыт другой документ. Ожидался: '{_pendingPersistencePath}', активен: '{activePath}'. Открой именно сохранённую тестовую копию.");
                Log("Путь совпадает с документом, в котором проводился тест. Считываем исполнения повторно.");
                ReadManager(manager);
            });

            if (!connected)
                throw new InvalidOperationException("Не удалось подключиться к повторно открытому документу.");

            var actual = _rows.Select(ToSnapshot).ToList();
            bool persistedStateMatches = AreMultisetsEqual(_expectedPersistedRows, actual);
            string diff = DescribeMultisetDifference(_expectedPersistedRows, actual);
            Log($"Проверка сохранённого состояния: PersistedStateMatches={persistedStateMatches}; ExpectedRows={_expectedPersistedRows.Count}; ActualRows={actual.Count}.");

            if (persistedStateMatches)
            {
                _pendingPersistenceCheck = false;
                Log("УСПЕХ: после ручного сохранения, закрытия и повторного открытия все базовые обозначения, номера исполнений и остальные строки совпали с ожидаемым состоянием.");
                StatusText.Text = "СОХРАНЕНИЕ ПОДТВЕРЖДЕНО";
                MessageBox.Show("Отлично! Сохранение подтверждено. После закрытия и повторного открытия документа всё семейство осталось в ожидаемом состоянии, номера исполнений на месте, остальные строки совпадают.",
                    "Сохранение подтверждено", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                LogError("СОХРАНЁННОЕ СОСТОЯНИЕ НЕ СОВПАЛО", new InvalidOperationException(diff));
                StatusText.Text = "Сохранённое состояние отличается — см. лог";
                MessageBox.Show("После повторного открытия состав обозначений отличается от ожидаемого. Не используй этот файл как доказательство корректности; сохрани лог и отправь его мне.\n\n" + diff,
                    "Проверка сохранения", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        catch (Exception ex)
        {
            LogError("Проверка после повторного открытия прервана", ex);
            StatusText.Text = "Ошибка проверки сохранения — см. лог";
            MessageBox.Show("Не удалось завершить проверку повторно открытого файла. Сохрани лог и отправь его мне.\n\n" + ex.Message,
                "Проверка сохранения", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            RefreshPersistenceButtonState();
        }
    }

    private List<ExecutionSnapshot> ReadSnapshotsWithoutGrid(object manager, int count)
    {
        var snapshots = new List<ExecutionSnapshot>();
        for (int index = 0; index < count; index++)
        {
            string baseMark = TryGetMarking(manager, index, 1, "Базовая часть при повторной проверке перед записью");
            string number = TryGetMarking(manager, index, 2, "Номер при повторной проверке перед записью");
            string full = TryGetMarking(manager, index, -1, "Полное обозначение при повторной проверке перед записью");
            string variant = TryGetVariant(manager, index);
            if (baseMark == "<ошибка>" || number == "<ошибка>" || full == "<ошибка>" || variant == "<ошибка>")
                throw new InvalidOperationException($"Не удалось заново прочитать строку исполнения с индексом {index}. Запись отменена.");
            snapshots.Add(new ExecutionSnapshot(baseMark, number, full, variant));
        }
        return snapshots;
    }

    private static ExecutionSnapshot ToSnapshot(ExecutionRow row)
        => new(row.BaseMarking, row.ExecutionNumber, row.FullMarking, row.IsVariant);

    private static Dictionary<ExecutionSnapshot, int> ToMultiset(IEnumerable<ExecutionSnapshot> rows)
    {
        var result = new Dictionary<ExecutionSnapshot, int>();
        foreach (var row in rows)
        {
            result.TryGetValue(row, out int count);
            result[row] = count + 1;
        }
        return result;
    }

    private static bool AreMultisetsEqual(IEnumerable<ExecutionSnapshot> expected, IEnumerable<ExecutionSnapshot> actual)
    {
        var left = ToMultiset(expected);
        var right = ToMultiset(actual);
        return left.Count == right.Count && left.All(pair => right.TryGetValue(pair.Key, out int count) && count == pair.Value);
    }

    private static bool AreStringMultisetsEqual(IEnumerable<string> expected, IEnumerable<string> actual)
    {
        var left = expected.GroupBy(value => value, StringComparer.Ordinal).ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);
        var right = actual.GroupBy(value => value, StringComparer.Ordinal).ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);
        return left.Count == right.Count && left.All(pair => right.TryGetValue(pair.Key, out int count) && count == pair.Value);
    }

    private static string DescribeMultisetDifference(IEnumerable<ExecutionSnapshot> expected, IEnumerable<ExecutionSnapshot> actual)
    {
        var left = ToMultiset(expected);
        var right = ToMultiset(actual);
        var lines = new List<string>();
        foreach (var item in left.Keys.Union(right.Keys))
        {
            left.TryGetValue(item, out int expectedCount);
            right.TryGetValue(item, out int actualCount);
            if (expectedCount != actualCount)
                lines.Add($"'{item.FullMarking}' (Base='{item.BaseMarking}', Number='{item.ExecutionNumber}', Variant='{item.IsVariant}'): ожидалось {expectedCount}, фактически {actualCount}");
        }
        return lines.Count == 0 ? "Различий по составу строк не обнаружено." : string.Join(Environment.NewLine, lines);
    }

    private static bool PathsEqual(string first, string second)
    {
        if (string.IsNullOrWhiteSpace(first) || string.IsNullOrWhiteSpace(second)) return false;
        try { return string.Equals(Path.GetFullPath(first), Path.GetFullPath(second), StringComparison.OrdinalIgnoreCase); }
        catch { return string.Equals(first.Trim(), second.Trim(), StringComparison.OrdinalIgnoreCase); }
    }

    private static object? GetActiveComObject(Guid clsid)
    {
        IntPtr ptr = IntPtr.Zero;
        int hr = NativeCom.GetActiveObject(ref clsid, IntPtr.Zero, out ptr);
        if (hr < 0 || ptr == IntPtr.Zero) return null;
        try { return Marshal.GetObjectForIUnknown(ptr); }
        finally { Marshal.Release(ptr); }
    }

    private static void ReleaseCom(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
        {
            try { Marshal.ReleaseComObject(value); } catch { }
        }
    }

    private void Log(string text)
    {
        Dispatcher.Invoke(() =>
        {
            _log.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] {text}");
            LogBox.Text = _log.ToString();
            LogBox.ScrollToEnd();
        });
    }

    private void LogError(string title, Exception ex)
        => Log($"ОШИБКА — {title}: {ex.GetType().Name}; HRESULT=0x{ex.HResult:X8}; {ex.Message}");

    private void SaveLogButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new SaveFileDialog
        {
            Filter = "Текстовый лог (*.txt)|*.txt",
            FileName = $"EmbodimentProbe_{DateTime.Now:yyyyMMdd_HHmmss}.txt"
        };
        if (dialog.ShowDialog() == true)
        {
            File.WriteAllText(dialog.FileName, _log.ToString(), Encoding.UTF8);
            MessageBox.Show("Лог сохранён.", "Диагностика", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

    private void ClearButton_Click(object sender, RoutedEventArgs e)
    {
        if (_rows.Count > 0 && _hasSuccessfulRead)
        {
            var answer = MessageBox.Show("Очистить таблицу и сбросить сведения о последнем чтении?", "Очистка", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (answer != MessageBoxResult.Yes) return;
        }
        _rows.Clear();
        _log.Clear();
        LogBox.Clear();
        _lastReadDocumentPath = string.Empty;
        _hasSuccessfulRead = false;
        _pendingPersistenceCheck = false;
        _pendingPersistencePath = string.Empty;
        _expectedPersistedRows.Clear();
        PersistenceCheckBox.IsChecked = false;
        RefreshPersistenceButtonState();
        NewBaseMarkingBox.Clear();
        ConfirmationPhraseBox.Clear();
        TestCopyCheckBox.IsChecked = false;
        StatusText.Text = "Очищено";
        RefreshWriteButtonState();
    }

    private static class NativeCom
    {
        [DllImport("oleaut32.dll", PreserveSig = true)]
        public static extern int GetActiveObject(ref Guid rclsid, IntPtr pvReserved, out IntPtr ppunk);
    }
}

internal sealed record ExecutionSnapshot(string BaseMarking, string ExecutionNumber, string FullMarking, string IsVariant);

public sealed class ExecutionRow
{
    public int Index { get; set; }
    public string BaseMarking { get; set; } = "";
    public string ExecutionNumber { get; set; } = "";
    public string FullMarking { get; set; } = "";
    public string IsVariant { get; set; } = "";
}

internal static class TypeLibInterfaceCatalog
{
    public sealed record Result(Dictionary<string, Guid> InterfaceIds, string Report);

    public static Result FindInterfaces(Guid typeLibId, params string[] targetNames)
    {
        var ids = new Dictionary<string, Guid>(StringComparer.OrdinalIgnoreCase);
        var report = new StringBuilder();
        ITypeLib? typeLib = null;

        try
        {
            typeLib = LoadRegisteredTypeLibrary(typeLibId, report);
            int count = typeLib.GetTypeInfoCount();
            report.AppendLine($"TypeLib {typeLibId}: найдено typeinfo={count}.");

            for (int i = 0; i < count; i++)
            {
                ITypeInfo? info = null;
                IntPtr attrPtr = IntPtr.Zero;
                try
                {
                    typeLib.GetTypeInfo(i, out info);
                    info.GetTypeAttr(out attrPtr);
                    var attr = Marshal.PtrToStructure<TYPEATTR>(attrPtr);
                    info.GetDocumentation(-1, out string name, out _, out _, out _);
                    if (targetNames.Contains(name, StringComparer.OrdinalIgnoreCase))
                    {
                        ids[name] = attr.guid;
                        report.AppendLine($"Интерфейс {name}: IID={attr.guid}");
                    }
                }
                catch (Exception ex) { report.AppendLine($"TypeInfo #{i}: {ex.GetType().Name}: {ex.Message}"); }
                finally
                {
                    if (attrPtr != IntPtr.Zero && info is not null) info.ReleaseTypeAttr(attrPtr);
                    if (info is not null && Marshal.IsComObject(info)) Marshal.ReleaseComObject(info);
                }
            }

            if (targetNames.Any(name => !ids.ContainsKey(name)))
                report.AppendLine("Предупреждение: один или несколько интерфейсов не найдены в библиотеке типов API7.");
        }
        catch (Exception ex)
        {
            report.AppendLine($"Ошибка чтения TypeLib: {ex.GetType().Name}; HRESULT=0x{ex.HResult:X8}; {ex.Message}");
        }
        finally
        {
            if (typeLib is not null && Marshal.IsComObject(typeLib)) Marshal.ReleaseComObject(typeLib);
        }

        return new Result(ids, report.ToString());
    }

    private static ITypeLib LoadRegisteredTypeLibrary(Guid typeLibId, StringBuilder report)
    {
        string root = $@"TypeLib\{{{typeLibId}}}";
        using var key = Registry.ClassesRoot.OpenSubKey(root);
        if (key is null)
        {
            report.AppendLine($"В реестре отсутствует раздел {root}.");
            throw new InvalidOperationException($"В реестре отсутствует {root}.");
        }

        var versions = key.GetSubKeyNames()
            .Select(name => (Name: name, Parsed: Version.TryParse(name, out var v) ? v : null))
            .Where(item => item.Parsed is not null)
            .OrderByDescending(item => item.Parsed)
            .ToList();

        foreach (var version in versions)
        {
            if (version.Parsed!.Major < 0 || version.Parsed.Major > ushort.MaxValue ||
                version.Parsed.Minor < 0 || version.Parsed.Minor > ushort.MaxValue) continue;

            using var versionKey = key.OpenSubKey(version.Name);
            if (versionKey is null) continue;

            var localeIds = new HashSet<int> { 0 };
            foreach (string localeName in versionKey.GetSubKeyNames())
            {
                if (int.TryParse(localeName, NumberStyles.Integer, CultureInfo.InvariantCulture, out int decimalLcid)) localeIds.Add(decimalLcid);
                if (int.TryParse(localeName, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int hexLcid)) localeIds.Add(hexLcid);
            }

            foreach (int lcid in localeIds)
            {
                try
                {
                    Guid id = typeLibId;
                    NativeTypeLib.LoadRegTypeLib(ref id, (ushort)version.Parsed.Major, (ushort)version.Parsed.Minor, lcid, out ITypeLib registered);
                    report.AppendLine($"LoadRegTypeLib успешно: версия {version.Name}, LCID={lcid}.");
                    return registered;
                }
                catch (Exception ex)
                {
                    report.AppendLine($"LoadRegTypeLib: версия {version.Name}, LCID={lcid}: HRESULT=0x{ex.HResult:X8}; {ex.Message}");
                }
            }
        }

        foreach (var version in versions)
        {
            using var versionKey = key.OpenSubKey(version.Name);
            if (versionKey is null) continue;
            foreach (string locale in versionKey.GetSubKeyNames())
            {
                using var localeKey = versionKey.OpenSubKey(locale);
                if (localeKey is null) continue;
                foreach (string platform in new[] { "win64", "win32" })
                {
                    using var platformKey = localeKey.OpenSubKey(platform);
                    string? path = platformKey?.GetValue(null) as string;
                    if (string.IsNullOrWhiteSpace(path)) continue;
                    path = Environment.ExpandEnvironmentVariables(path.Trim().Trim('"'));
                    if (!File.Exists(path))
                    {
                        report.AppendLine($"TypeLib-файл не найден ({platform}): {path}");
                        continue;
                    }
                    try
                    {
                        NativeTypeLib.LoadTypeLibEx(path, RegKind.None, out ITypeLib loaded);
                        report.AppendLine($"LoadTypeLibEx успешно: {path}");
                        return loaded;
                    }
                    catch (Exception ex) { report.AppendLine($"LoadTypeLibEx не удался: {path}; HRESULT=0x{ex.HResult:X8}; {ex.Message}"); }
                }
            }
        }

        throw new InvalidOperationException("Не удалось загрузить TypeLib API7 ни через LoadRegTypeLib, ни из зарегистрированных файловых путей.");
    }

    private enum RegKind { None = 2 }

    private static class NativeTypeLib
    {
        [DllImport("oleaut32.dll", PreserveSig = false)]
        public static extern void LoadTypeLibEx(string file, RegKind regKind, out ITypeLib typeLib);

        [DllImport("oleaut32.dll", PreserveSig = false)]
        public static extern void LoadRegTypeLib(ref Guid guid, ushort major, ushort minor, int lcid, out ITypeLib typeLib);
    }
}
