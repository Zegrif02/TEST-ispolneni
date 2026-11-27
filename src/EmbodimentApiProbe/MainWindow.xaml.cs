using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Windows;

namespace EmbodimentApiProbe;

public partial class MainWindow : Window
{
    private readonly ObservableCollection<ExecutionRow> _rows = new();
    private readonly StringBuilder _log = new();
    private const string KompasProgId = "KOMPAS.Application.7";
    private static readonly Guid Api7TypeLibGuid = new("69AC2981-37C0-4379-84FD-5DD2F3C0A520");

    public MainWindow()
    {
        InitializeComponent();
        ResultsGrid.ItemsSource = _rows;
        Log("Диагностическая программа 0.1.0 запущена.");
        Log("Режим строго только чтение. Изменяющие COM-вызовы в программу не включены.");
    }

    private void ReadButton_Click(object sender, RoutedEventArgs e)
    {
        ReadButton.IsEnabled = false;
        StatusText.Text = "Выполняется...";
        _rows.Clear();
        Log("============================================================");
        Log("Начало диагностики активного документа.");

        try
        {
            // COM calls remain on the WPF STA thread. Do not move this scan to Task.Run.
            ReadActiveDocument();
            StatusText.Text = $"Готово. Строк: {_rows.Count}";
        }
        catch (Exception ex)
        {
            LogError("Необработанная ошибка", ex);
            StatusText.Text = "Ошибка — подробности в журнале";
        }
        finally
        {
            ReadButton.IsEnabled = true;
        }
    }

    private void ReadActiveDocument()
    {
        object? app = null;
        object? document = null;
        object? doc3d = null;
        object? manager = null;
        IntPtr appPtr = IntPtr.Zero;
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
                return;
            }

            app = GetActiveComObject(appType.GUID);
            if (app is null)
            {
                Log("Активный экземпляр КОМПАС-3D не найден. Новый экземпляр не запускался.");
                return;
            }

            Log("Подключение к существующему экземпляру успешно.");
            try
            {
                document = ((dynamic)app).ActiveDocument;
            }
            catch (Exception ex)
            {
                LogError("Получение ActiveDocument", ex);
                return;
            }

            if (document is null)
            {
                Log("В КОМПАС нет активного документа.");
                return;
            }

            LogDocumentInfo(document);
            Log("Загрузка зарегистрированной библиотеки типов API7...");
            var interfaces = TypeLibInterfaceCatalog.FindInterfaces(Api7TypeLibGuid, "IEmbodimentsManager", "IKompasDocument3D");
            Log(interfaces.Report);

            if (!interfaces.InterfaceIds.TryGetValue("IKompasDocument3D", out var document3dIid))
            {
                Log("Не удалось найти IID интерфейса IKompasDocument3D в зарегистрированной библиотеке типов API7.");
                return;
            }
            if (!interfaces.InterfaceIds.TryGetValue("IEmbodimentsManager", out var managerIid))
            {
                Log("Не удалось найти IID IEmbodimentsManager в зарегистрированной библиотеке типов API7.");
                return;
            }

            docPtr = Marshal.GetIUnknownForObject(document);
            if (Marshal.QueryInterface(docPtr, ref document3dIid, out doc3dPtr) < 0 || doc3dPtr == IntPtr.Zero)
            {
                Log("QueryInterface IKompasDocument3D не выполнен. Пробуем QueryInterface IEmbodimentsManager напрямую от ActiveDocument...");
                if (Marshal.QueryInterface(docPtr, ref managerIid, out managerPtr) < 0 || managerPtr == IntPtr.Zero)
                {
                    Log("Не удалось получить IEmbodimentsManager от ActiveDocument.");
                    return;
                }
            }
            else
            {
                doc3d = Marshal.GetObjectForIUnknown(doc3dPtr);
                Log("Интерфейс IKompasDocument3D получен через QueryInterface.");
                var basePtr = doc3dPtr;
                if (Marshal.QueryInterface(basePtr, ref managerIid, out managerPtr) < 0 || managerPtr == IntPtr.Zero)
                {
                    Log("QueryInterface IEmbodimentsManager от IKompasDocument3D не выполнен. Пробуем ActiveDocument...");
                    if (Marshal.QueryInterface(docPtr, ref managerIid, out managerPtr) < 0 || managerPtr == IntPtr.Zero)
                    {
                        Log("Не удалось получить IEmbodimentsManager ни от IKompasDocument3D, ни от ActiveDocument.");
                        return;
                    }
                }
            }

            manager = Marshal.GetObjectForIUnknown(managerPtr);
            Log("IEmbodimentsManager получен. Начинаем только чтение.");
            ReadManager(manager);
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
            if (appPtr != IntPtr.Zero) Marshal.Release(appPtr);
        }
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
        Log($"Количество исполнений: {(count.HasValue ? count.Value.ToString() : "не удалось получить")}");

        TryInvokeAndLog(manager, "GetEmbodimentsTree", "Дерево исполнений", Array.Empty<object>());
        TryInvokeAndLog(manager, "GetCurrentEmbodimentIndex", "Индекс текущего исполнения", Array.Empty<object>());
        TryInvokeAndLog(manager, "GetCurrentEmbodimentMarking", "Обозначение текущего исполнения (полное)", new object[] { -1, false });

        if (!count.HasValue || count.Value <= 0)
        {
            Log("Количество исполнений не получено или равно нулю. Автоматический перебор индексов отменён.");
            return;
        }

        Log("--- Чтение обозначений по индексам (попытка 0-based) ---");
        for (int index = 0; index < count.Value; index++)
        {
            ReadOneIndex(manager, index);
        }

        Log("--- Дополнительная проверка 1-based индексации (только чтение; последний индекс count) ---");
        ReadOneIndex(manager, count.Value);
    }

    private void ReadOneIndex(object manager, int index)
    {
        string baseMark = TryGetMarking(manager, index, 1, "Базовая часть");
        string number = TryGetMarking(manager, index, 2, "Номер исполнения");
        string full = TryGetMarking(manager, index, -1, "Полное обозначение");
        string variant = TryGetVariant(manager, index);

        Log($"Индекс {index}: Base='{baseMark}' | Number='{number}' | Full='{full}' | IsVariant='{variant}'");

        // Добавляем строки, если хоть один запрос для индекса был успешным.
        if (baseMark != "<ошибка>" || number != "<ошибка>" || full != "<ошибка>" || variant != "<ошибка>")
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
                if (value is not null && int.TryParse(Convert.ToString(value), out int count))
                {
                    Log($"Получено через {property}: {count}");
                    return count;
                }
            }
            catch (Exception ex)
            {
                LogError($"Чтение {property}", ex);
            }
        }
        return null;
    }

    private string TryGetMarking(object manager, int index, int type, string label)
    {
        try
        {
            dynamic d = manager;
            object? result = d.GetEmbodimentMarking(index, type, false);
            return Convert.ToString(result) ?? string.Empty;
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
            dynamic d = manager;
            object? result = d.IsVariant(index);
            return Convert.ToString(result) ?? "";
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
        catch (Exception ex)
        {
            LogError(label, ex);
        }
    }

    private static string FormatComResult(object? value)
    {
        if (value is null) return "<null>";
        if (value is Array array)
        {
            var parts = new List<string>();
            foreach (object? item in array) parts.Add(Convert.ToString(item) ?? "<null>");
            return $"Array[{array.Length}]: " + string.Join(" | ", parts);
        }
        return Convert.ToString(value) ?? value.ToString() ?? "<unknown>";
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
            Log($"{propertyName}: {Convert.ToString(value) ?? "<null>"}");
        }
        catch (Exception ex)
        {
            LogError($"Свойство {propertyName}", ex);
        }
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
    {
        Log($"ОШИБКА — {title}: {ex.GetType().Name}; HRESULT=0x{ex.HResult:X8}; {ex.Message}");
    }

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
        _rows.Clear();
        _log.Clear();
        LogBox.Clear();
        StatusText.Text = "Очищено";
    }

    private static class NativeCom
    {
        [DllImport("oleaut32.dll", PreserveSig = true)]
        public static extern int GetActiveObject(ref Guid rclsid, IntPtr pvReserved, out IntPtr ppunk);
    }
}

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
            typeLib = LoadRegisteredTypeLibrary(typeLibId);
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
                catch (Exception ex)
                {
                    report.AppendLine($"TypeInfo #{i}: {ex.GetType().Name}: {ex.Message}");
                }
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

    private static ITypeLib LoadRegisteredTypeLibrary(Guid typeLibId)
    {
        // Сначала пробуем загрузить TypeLib из зарегистрированного пути в реестре.
        string root = $@"TypeLib\{{{typeLibId}}}";
        using var key = Registry.ClassesRoot.OpenSubKey(root);
        if (key is null) throw new InvalidOperationException($"В реестре отсутствует {root}.");

        foreach (string version in key.GetSubKeyNames().OrderByDescending(s => s, StringComparer.OrdinalIgnoreCase))
        {
            using var versionKey = key.OpenSubKey(version);
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
                    path = Environment.ExpandEnvironmentVariables(path.Trim('"'));
                    if (File.Exists(path))
                    {
                        NativeTypeLib.LoadTypeLibEx(path, RegKind.None, out ITypeLib loaded);
                        return loaded;
                    }
                }
            }
        }

        // Если typelib зарегистрирована как встроенная/без файла, пробуем LoadRegTypeLib.
        NativeTypeLib.LoadRegTypeLib(ref typeLibId, 1, 0, 0, out ITypeLib result);
        return result;
    }

    private enum RegKind
    {
        None = 2
    }

    private static class NativeTypeLib
    {
        [DllImport("oleaut32.dll", PreserveSig = false)]
        public static extern void LoadTypeLibEx(string file, RegKind regKind, out ITypeLib typeLib);

        [DllImport("oleaut32.dll", PreserveSig = false)]
        public static extern void LoadRegTypeLib(ref Guid guid, ushort major, ushort minor, int lcid, out ITypeLib typeLib);
    }
}
