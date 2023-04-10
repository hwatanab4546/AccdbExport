using Reactive.Bindings;
using Reactive.Bindings.Disposables;
using Reactive.Bindings.Extensions;
using Reactive.Bindings.TinyLinq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using XAccess;

namespace AccdbExport
{
    internal class MainWindowViewModel : INotifyPropertyChanged, IDisposable
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public ReactivePropertySlim<string> AccdbPath { get; } = new ReactivePropertySlim<string>();
        public ReactiveCommandSlim RunCommand { get; }
        public ReactivePropertySlim<bool> IsBusy { get; } = new ReactivePropertySlim<bool>();
        public ReactivePropertySlim<string> Status { get; } = new ReactivePropertySlim<string>();

        public MainWindowViewModel()
        {
            RunCommand
                = new[]
                {
                    AccdbPath.Select(s => !string.IsNullOrEmpty(s) && System.IO.File.Exists(s)),
                    IsBusy.Select(s => !s),

                }.CombineLatest(s => s[0] && s[1])
                .ToReactiveCommandSlim()
                .WithSubscribe(async () =>
                {
                    IsBusy.Value = true;
                    try
                    {
                        await Task.Run(() =>
                        {
                            AccdbExport(AccdbPath.Value, ReportStatus);
                        });
                    }
                    finally
                    {
                        IsBusy.Value = false;
                    }
                })
                .AddTo(disposables);
        }

        private void ReportStatus(string status)
        {
            Application.Current.Dispatcher.BeginInvoke(new Action<string>(s => Status.Value = s), status);
        }

        private void AccdbExport(string accdbPath, Action<string> reportStatus = null)
        {
            try
            {
                var folder = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(accdbPath), System.IO.Path.GetFileNameWithoutExtension(accdbPath) + "_contents");

                if (System.IO.Directory.Exists(folder))
                {
                    if (MessageBox.Show("出力先フォルダが既に存在しています。内容を破棄してよろしいですか？", "確認", MessageBoxButton.OKCancel, MessageBoxImage.Question) != MessageBoxResult.OK)
                    {
                        return;
                    }
                    System.IO.Directory.Delete(folder, true);
                }
                System.IO.Directory.CreateDirectory(folder);

                var timeList = new List<(string Title, long Time)> { ("(dummy)", 0), };
                var stopwatch = new System.Diagnostics.Stopwatch();
                stopwatch.Start();

                reportStatus?.Invoke("MS Accessを起動しています... (1/5)");

                using (var application = new XApplication(accdbPath))
                {
                    timeList.Add(("Starting MS Access", stopwatch.ElapsedMilliseconds));

                    // VBAコードの出力
                    foreach (var vBComponent in application.VBComponents)
                    {
                        string ext;
                        switch (vBComponent.Type)
                        {
                            case Xvbext_ComponentType.vbext_ct_Document:
                            case Xvbext_ComponentType.vbext_ct_StdModule:
                                ext = ".bas";
                                break;
                            case Xvbext_ComponentType.vbext_ct_ClassModule:
                                ext = ".cls";
                                break;
#if false
                            // この方法ではForm情報は抽出できない模様(mdb用？)
                            case Xvbext_ComponentType.vbext_ct_MSForm:
                                ext = ".frm";
                                break;
#endif
                            default:
                                continue;
                        }

                        var filePath = System.IO.Path.Combine(folder, $"{vBComponent.Name}{ext}");

                        vBComponent.Export(filePath);
                    }

                    timeList.Add(("Exporting VBAs", stopwatch.ElapsedMilliseconds));

                    reportStatus?.Invoke("フォーム情報を出力しています... (2/5)");

                    // フォーム情報の出力
                    foreach (var form in application.AllForms)
                    {
                        var filepath = System.IO.Path.Combine(folder, $"{form.Name}.frm");

                        application.SaveAsText(XAcObjectType.acForm, form.Name, filepath);
                    }

                    timeList.Add(("Exporting Forms", stopwatch.ElapsedMilliseconds));

                    reportStatus?.Invoke("テーブル定義を出力しています... (3/5)");

                    using (var db = application.CurrentDb())
                    {
                        // テーブル定義の出力
                        foreach (var tabledef in db.TableDefs)
                        {
                            if (tabledef.Name.StartsWith("MSys"))
                            {
                                continue;
                            }

                            var filepath = System.IO.Path.Combine(folder, $"{tabledef.Name}.xsd");

                            application.ExportXML(XAcExportXMLObjectType.acExportTable, tabledef.Name, filepath);
                        }

                        timeList.Add(("Exporting TableDefs", stopwatch.ElapsedMilliseconds));

                        reportStatus?.Invoke("クエリを出力しています... (4/5)");

                        // クエリの出力
                        foreach (var querydef in db.QueryDefs)
                        {
                            if (querydef.Name.StartsWith("~"))
                            {
                                continue;
                            }

                            var filepath = System.IO.Path.Combine(folder, $"{querydef.Name}.sql");

                            using (var sw = new System.IO.StreamWriter(filepath, false))
                            {
                                sw.Write(querydef.SQL);
                            }
                        }

                        timeList.Add(("Exporting QueryDefs", stopwatch.ElapsedMilliseconds));
                    }

                    reportStatus?.Invoke("MS Accessを終了しています... (5/5)");

                    application.Quit();

                    timeList.Add(("Quitting MS Access", stopwatch.ElapsedMilliseconds));

                    reportStatus?.Invoke("終了しています...");
                }

                stopwatch.Stop();
                timeList.Add(("Releasing Application object(?)", stopwatch.ElapsedMilliseconds));

                MessageBox.Show(string.Join(Environment.NewLine, timeList.Take(timeList.Count - 1).Zip(timeList.Skip(1), (a, b) => $"{b.Title}: {b.Time - a.Time}[ms]")), "完了しました");

                reportStatus?.Invoke("");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "例外発生", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private readonly CompositeDisposable disposables = new CompositeDisposable();
        public void Dispose() => disposables.Dispose();
    }
}
