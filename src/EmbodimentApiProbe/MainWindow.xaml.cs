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

    public MainWindow()
    {
        InitializeComponent();
        ResultsGrid.ItemsSource = _rows;
        Log("Диагностическая программа 0.2.1 запущена.");
        Log("Режим чтения доступен всегда. Тест записи и одного перестроения отдельно защищён подтверждениями и требует тестовую копию.");
        Log("Программа не вызывает Save КОМПАС. После тестовой записи выполняется ровно один IKompasDocument3D.RebuildDocument().");
    }

    private void ReadButton_Click(object sender, RoutedEventArgs e)
    {
        ReadButton.IsEnabled = false;
        StatusText.Text = "Выполняется...";
        _rows.Clear();
        _lastReadDocumentPath = string.Empty;
        _hasSuccessfulRead = false;
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
            MessageBox.Show("Для первого теста варианты геометрии не поддерживаются. Выбери обычное исполнение.", "Тест записи", MessageBoxButton.OK, MessageBoxImage.Warning);
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

        string expectedFull = newBaseMarking + selected.ExecutionNumber;
        string prompt = "СЕЙЧАС БУДЕТ ИЗМЕНЕНО ОБОЗНАЧЕНИЕ В ОТКРЫТОМ ДОКУМЕНТЕ КОМПАС.\n\n" +
                        $"Файл: {_lastReadDocumentPath}\n" +
                        $"Индекс: {selected.Index}\n" +
                        $"Было: {selected.FullMarking}\n" +
                        $"Новая база: {newBaseMarking}\n" +
                        $"Номер исполнения должен остаться: {selected.ExecutionNumber}\n" +
                        $"Ожидаемое полное обозначение: {expectedFull}\n\n" +
                        "Программа вызовет только SetEmbodimentMarking(index, ksVMBaseMarking, newBase). " +
                        "После SetEmbodimentMarking программа ОДИН РАЗ вызовет IKompasDocument3D.RebuildDocument(), затем повторно подключится и проверит все исполнения. Save НЕ вызывается.\n\n" +
                        "Перестроение может затронуть документ. Продолжать только на отдельной одноразовой копии!";

        if (MessageBox.Show(prompt, "Последнее подтверждение изменения", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No) != MessageBoxResult.Yes)
        {
            Log("Пользователь отменил тест записи на последнем подтверждении.");
            return;
        }

        WriteTestButton.IsEnabled = false;
        StatusText.Text = "Выполняется контролируемая запись...";
        Log("============================================================");
        Log("НАЧАЛО КОНТРОЛИРУЕМОГО ТЕСТА ЗАПИСИ И ОДНОГО ПЕРЕСТРОЕНИЯ. Save не вызывается.");

        // Сохраняем исходное состояние всех строк, чтобы после перестроения
        // проверить не только выбранное исполнение, но и отсутствие побочных изменений.
        var beforeRows = _rows.ToDictionary(
            row => row.Index,
            row => new ExecutionSnapshot(row.BaseMarking, row.ExecutionNumber, row.FullMarking, row.IsVariant));
        int selectedIndex = selected.Index;
        string selectedBaseBefore = selected.BaseMarking;
        string selectedNumberBefore = selected.ExecutionNumber;
        string selectedFullBefore = selected.FullMarking;

        bool rebuildSucceeded = false;
        Exception? rebuildException = null;
        bool writeAccepted = false;

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
                if (beforeBase != selectedBaseBefore || beforeNumber != selectedNumberBefore || beforeFull != selectedFullBefore)
                    throw new InvalidOperationException("Данные выбранного исполнения изменились после последнего чтения. Запись отменена — нажми считывание ещё раз.");

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
                        Log("ПРЕДУПРЕЖДЕНИЕ: RebuildDocument вернул FALSE. Всё равно выполняем отдельное повторное чтение, чтобы зафиксировать фактическое состояние.");
                }
                catch (Exception ex)
                {
                    rebuildException = ex;
                    LogError("Вызов IKompasDocument3D.RebuildDocument", ex);
                }
            });

            if (!writeConnected)
                throw new InvalidOperationException("Не удалось подключиться к активному КОМПАС для записи. Перестроение и повторная проверка не выполнены.");

            // Закрываем предыдущие COM-ссылки и открываем свежие через ROT.
            // Так мы читаем состояние после перестроения не из старого менеджера.
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

            ExecutionRow? refreshedTarget = _rows.FirstOrDefault(row => row.Index == selectedIndex);
            if (refreshedTarget is null)
                throw new InvalidOperationException($"После перестроения не удалось повторно прочитать исполнение с индексом {selectedIndex}. Не сохраняй тестовую копию.");

            bool baseMatches = string.Equals(refreshedTarget.BaseMarking, newBaseMarking, StringComparison.Ordinal);
            bool numberPreserved = string.Equals(refreshedTarget.ExecutionNumber, selectedNumberBefore, StringComparison.Ordinal);
            bool fullMatches = string.Equals(refreshedTarget.FullMarking, expectedFull, StringComparison.Ordinal);

            int unchangedOthers = 0;
            var collateralChanges = new List<string>();
            foreach (var oldRow in beforeRows)
            {
                if (oldRow.Key == selectedIndex) continue;
                ExecutionRow? current = _rows.FirstOrDefault(row => row.Index == oldRow.Key);
                if (current is null)
                {
                    collateralChanges.Add($"индекс {oldRow.Key}: строка отсутствует после перестроения");
                    continue;
                }

                bool unchanged = string.Equals(current.BaseMarking, oldRow.Value.BaseMarking, StringComparison.Ordinal) &&
                                 string.Equals(current.ExecutionNumber, oldRow.Value.ExecutionNumber, StringComparison.Ordinal) &&
                                 string.Equals(current.FullMarking, oldRow.Value.FullMarking, StringComparison.Ordinal) &&
                                 string.Equals(current.IsVariant, oldRow.Value.IsVariant, StringComparison.Ordinal);
                if (unchanged) unchangedOthers++;
                else collateralChanges.Add($"индекс {oldRow.Key}: было '{oldRow.Value.FullMarking}', стало '{current.FullMarking}'");
            }

            Log("--- Итоговая проверка после одного RebuildDocument и повторного подключения ---");
            Log($"Целевое исполнение {selectedIndex}: Base='{refreshedTarget.BaseMarking}' | Number='{refreshedTarget.ExecutionNumber}' | Full='{refreshedTarget.FullMarking}'");
            Log($"Проверки: BaseMatches={baseMatches}; ExecutionNumberPreserved={numberPreserved}; FullMatches={fullMatches}; RebuildSucceeded={rebuildSucceeded}; UnchangedOtherExecutions={unchangedOthers}/{Math.Max(0, beforeRows.Count - 1)}");

            bool othersUnchanged = collateralChanges.Count == 0 && unchangedOthers == Math.Max(0, beforeRows.Count - 1);
            if (baseMatches && numberPreserved && fullMatches && rebuildSucceeded && othersUnchanged && rebuildException is null)
            {
                Log("УСПЕХ: одного RebuildDocument достаточно. Базовая часть изменилась, номер исполнения сохранён, остальные исполнения не затронуты.");
                Log("Документ остался изменённым в памяти КОМПАС. Save не вызывался. Закрой тестовую копию без сохранения после проверки.");
                StatusText.Text = "Успех: одного перестроения достаточно — НЕ СОХРАНЯТЬ";
                MessageBox.Show("Тест пройден! Одного RebuildDocument оказалось достаточно: базовая часть обновилась, номер исполнения сохранился, остальные исполнения не изменились.\n\nПрограмма не сохраняла документ. Закрой тестовую копию без сохранения.",
                    "Тест перестроения завершён", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                string details = $"BaseMatches={baseMatches}; ExecutionNumberPreserved={numberPreserved}; FullMatches={fullMatches}; RebuildSucceeded={rebuildSucceeded}; UnchangedOtherExecutions={unchangedOthers}. " +
                                 $"После перестроения: Base='{refreshedTarget.BaseMarking}', Number='{refreshedTarget.ExecutionNumber}', Full='{refreshedTarget.FullMarking}'.";
                if (collateralChanges.Count > 0)
                    details += " Побочные изменения: " + string.Join("; ", collateralChanges);
                if (rebuildException is not null)
                    details += $" Ошибка RebuildDocument: {rebuildException.Message}";
                LogError("ИТОГОВАЯ ПРОВЕРКА НЕ ПРОЙДЕНА", new InvalidOperationException(details + " Не сохраняй тестовую копию; закрой её без сохранения."));
                StatusText.Text = "Проверка не пройдена — НЕ СОХРАНЯТЬ документ";
                MessageBox.Show("Проверка после одного перестроения не пройдена полностью. НЕ СОХРАНЯЙ документ; закрой тестовую копию без сохранения.\n\n" + details,
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
