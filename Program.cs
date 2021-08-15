using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenCvSharp;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Threading;
using System.Runtime.InteropServices;

namespace dorareko
{
    class Program
    {
        private static System.Threading.Timer timer;

        static void Main(string[] args)
        {
            const string imageFormat = ".jpg";

            int id = 0;
            var frameSize = new Size(640, 480);
            int fps = 0;
            int period = 0;
            double speedTolerance = 0;
            TimeSpan spanMin = new TimeSpan(0);
            TimeSpan spanMax = new TimeSpan(0);
            int minutes = 0;
            int seconds = 0;
            TimeSpan recordLength = new TimeSpan(0);
            string workdir = "";
            string savedir = "";
            double spaceTolerance = 0;
            string keyword = "";

            bool isRunning = false;
            bool isTargetWindowExist = false;

            // 設定ファイル読み込み
            try
            {
                var dic = ReadIni(@"setting.ini");
                id = Convert.ToInt32(dic["camera"]["id"]);
                fps = Convert.ToInt32(dic["camera"]["fps"]);
                period = 1000 / fps;
                speedTolerance = Convert.ToDouble(dic["camera"]["tolerance"]);
                spanMax = new TimeSpan(0, 0, 0, 0, (int) Math.Round(1000 / fps * speedTolerance));
                spanMin = new TimeSpan(0, 0, 0, 0, (int) Math.Round(1000 / fps / speedTolerance));
                minutes = Convert.ToInt32(dic["camera"]["minutes"]);
                seconds = Convert.ToInt32(dic["camera"]["seconds"]);
                recordLength = new TimeSpan(0, minutes, seconds);
                
                workdir = dic["storage"]["workdir"];
                savedir = dic["storage"]["savedir"];
                spaceTolerance = Convert.ToDouble(dic["storage"]["tolerance"]);

                keyword = dic["trigger"]["keyword"];
            }

            catch
            {
                MessageBox.Show("setting.iniの読み込みに失敗しました", "ドラレコ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(-1);
            }

            // RAMDISKのマウント待ち
            DateTime waitStart = DateTime.Now;
            TimeSpan waitLength = new TimeSpan(0, 0, 10);   // 10s

            while (true)
            {
                System.Threading.Thread.Sleep(100);

                System.IO.DriveInfo drive = new System.IO.DriveInfo(workdir.Substring(0, 1));
                if (drive.IsReady == true)
                {
                    break;
                }
                
                if (DateTime.Now - waitStart > waitLength)
                {
                    MessageBox.Show("作業フォルダのドライブのマウント待ちがタイムアウトしました", "ドラレコ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Environment.Exit(-1);
                }
            }

            // 作業フォルダ削除
            if (System.IO.Directory.Exists(workdir))
            {
                try
                {
                    System.IO.Directory.Delete(workdir, true);
                }
                catch
                {
                    MessageBox.Show("作業フォルダの削除に失敗しました", "ドラレコ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Environment.Exit(-1);
                }
            }

            // 作業フォルダ作成
            try
            {
                System.IO.Directory.CreateDirectory(workdir);
            }
            catch
            {
                MessageBox.Show("作業フォルダの作成に失敗しました", "ドラレコ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(-1);
            }

            // 保存先フォルダ用意
            if (System.IO.Directory.Exists(savedir) == false)
            {
                try
                {
                    System.IO.Directory.CreateDirectory(savedir);
                }
                catch
                {
                    MessageBox.Show("保存先フォルダの作成に失敗しました", "ドラレコ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Environment.Exit(-1);
                }
            }

            //// トリガをクリア
            //if (System.IO.File.Exists(trigpath))
            //{
            //    try
            //    {
            //        System.IO.File.Delete(trigpath);
            //    }
            //    catch
            //    {
            //        MessageBox.Show("トリガファイルの削除に失敗しました", "ドラレコ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //        Environment.Exit(-1);
            //    }
            //}

            // カメラ動作チェック
            var capture = new VideoCapture(id);
            if (!capture.IsOpened())
            {
                MessageBox.Show("指定IDのカメラに接続できませんでした", "ドラレコ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(-1);
            }
            frameSize = new Size(capture.FrameWidth, capture.FrameHeight);

            // カメラから作業フォルダに静止画を保存するチェック
            try
            {
                for (int i = 0; i < 10; i++)
                {
                    Mat mat = new Mat();
                    capture.Read(mat);
                    Cv2.PutText(mat, DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff"), new Point(10, 20), HersheyFonts.HersheyComplexSmall, 1, new Scalar(0, 255, 0), 1, LineTypes.AntiAlias);
                    mat.SaveImage(workdir + @"\" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + imageFormat);
                    System.Threading.Thread.Sleep(100);
                }
            }
            catch
            {
                MessageBox.Show("作業フォルダへの静止画の保存に失敗しました", "ドラレコ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(-1);
            }

            // ここで保存が完了しているのか検証不十分では？Sleepを減らしたり枚数を増やすとどうなる？

            // 作業フォルダの静止画から、保存先フォルダに動画を生成して、用済みの静止画を消去する
            try
            {
                string testpath = savedir + @"\" + "test.avi";

                if (System.IO.File.Exists(testpath)){
                    System.IO.File.Delete(testpath);
                }

                IEnumerable<string> files = System.IO.Directory.EnumerateFiles(workdir, "*" + imageFormat);
                var writer = new VideoWriter(@"C:\log\test.avi", FourCC.DIVX, fps, frameSize);
                foreach (string f in files)
                {
                    Mat mat = new Mat();
                    mat = Cv2.ImRead(f);
                    writer.Write(mat);
                    System.IO.File.Delete(f);
                }
                writer.Release();
                System.IO.File.Delete(testpath);
            }
            catch
            {
                MessageBox.Show("静止画から動画の生成または保存に失敗しました", "ドラレコ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(-1);
            }

            // タイマーで処理される内容（マルチスレッド）
            TimerCallback callback = state =>
            {

                // 静止画を保存する
                Mat mat = new Mat();
                capture.Read(mat);
                Cv2.PutText(mat, DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff"), new Point(10, 20), HersheyFonts.HersheyComplexSmall, 1, new Scalar(0, 255, 0), 1, LineTypes.AntiAlias);
                mat.SaveImage(workdir + @"\" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + imageFormat);

                // 古いファイルは順次削除する
                DateTime filetime = new DateTime(0);
                IEnumerable<string> files = System.IO.Directory.EnumerateFiles(workdir, "*" + imageFormat);

                foreach (string f in files)
                {
                    // ファイル名からDateTimeを取得
                    try
                    {
                        string t = Convert.ToString(System.IO.Path.GetFileNameWithoutExtension(f));
                        filetime = new DateTime(
                            Convert.ToInt32(t.Substring(0, 4)),
                            Convert.ToInt32(t.Substring(4, 2)),
                            Convert.ToInt32(t.Substring(6, 2)),
                            Convert.ToInt32(t.Substring(8, 2)),
                            Convert.ToInt32(t.Substring(10, 2)),
                            Convert.ToInt32(t.Substring(12, 2)),
                            Convert.ToInt32(t.Substring(14, 3)));
                    }
                    catch
                    {
                        MessageBox.Show("静止画のファイルに名前が異常なものがあります", "ドラレコ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Environment.Exit(-1);
                    }

                    // 空き容量チェック（10sec分以下になったら強制削除）
                    bool isDriveFull = false;
                    System.IO.DriveInfo drive = new System.IO.DriveInfo(files.ElementAt(0).Substring(0, 1));
                    System.IO.FileInfo fileInfo = new System.IO.FileInfo(files.ElementAt(0));
                    if (drive.AvailableFreeSpace < fileInfo.Length * 10 * fps)
                    {
                        isDriveFull = true;
                    }

                    // 録画時間外の古いファイルは削除する
                    // 空き容量がないときは古いものから順に削除する
                    if (DateTime.Now - filetime > recordLength || isDriveFull == true)
                    {
                        try
                        {
                            System.IO.File.Delete(f);   // すでに別スレッドで削除されているときはエラーになるので無視する
                        }
                        catch
                        {
                            // nothing
                        }
                    }
                    else
                    {
                        break;  // ファイル名＝時刻で古いものから列挙されているので、条件を満たしたら削除を打ち切る
                    }
                }
            };

            // テスト実行（10秒）
            DateTime testStart = DateTime.Now;
            TimeSpan testLength = new TimeSpan(0, 0, 10);   // 10s

            // タイマー開始
            timer = new System.Threading.Timer(callback, null, 0, period);

            while (true)
            {
                System.Threading.Thread.Sleep(100);

                if (DateTime.Now - testStart > testLength)
                {
                    // 実行中のスレッドがすべて終了するのを待つために、確実なDisposeを行う
                    // https://so-zou.jp/software/tech/programming/c-sharp/thread/
                    // timer.Dispose();の代わりに以下のようにする
                    using (System.Threading.WaitHandle waitHandle = new System.Threading.ManualResetEvent(false))
                    {
                        if (timer.Dispose(waitHandle))
                        {
                            const int millisecondsTimeout = 1000;
                            if (!waitHandle.WaitOne(millisecondsTimeout))
                            {
                                // 指定時間内に破棄されなかった
                            }
                        }
                    }

                    // 作業フォルダの静止画リスト
                    IEnumerable<string> files = System.IO.Directory.EnumerateFiles(workdir, "*" + imageFormat);

                    // フレームレートのテスト
                    List<DateTime> timeList = new List<DateTime>();
                    List<TimeSpan> spanList = new List<TimeSpan>();

                    foreach (string f in files)
                    {
                        string t = Convert.ToString(System.IO.Path.GetFileNameWithoutExtension(f));
                        DateTime d = new DateTime(
                            Convert.ToInt32(t.Substring(0, 4)),
                            Convert.ToInt32(t.Substring(4, 2)),
                            Convert.ToInt32(t.Substring(6, 2)),
                            Convert.ToInt32(t.Substring(8, 2)),
                            Convert.ToInt32(t.Substring(10, 2)),
                            Convert.ToInt32(t.Substring(12, 2)),
                            Convert.ToInt32(t.Substring(14, 3)));
                        timeList.Add(d);
                    }
                    for(int i = 1; i < timeList.Count(); i++)
                    {
                        spanList.Add(timeList[i] - timeList[i - 1]);
                    }

                    if (spanList.Min() < spanMin || spanList.Max() > spanMax)
                    {
                        MessageBox.Show("指定されたフレームレートが達成できませんでした", "ドラレコ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Environment.Exit(-1);
                    }

                    // 空き容量チェック
                    System.IO.DriveInfo drive = new System.IO.DriveInfo(files.ElementAt(0).Substring(0, 1));
                    System.IO.FileInfo fileInfo = new System.IO.FileInfo(files.ElementAt(0));
                    if (drive.AvailableFreeSpace < fileInfo.Length * (minutes * 60 + seconds) * fps * spaceTolerance)
                    {
                        MessageBox.Show("作業フォルダの空き容量が不足しています", "ドラレコ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Environment.Exit(-1);
                    }

                    // 動画にして保存する
                    try
                    {
                        string savepath = savedir + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".avi";

                        if (System.IO.File.Exists(savepath))
                        {
                            System.IO.File.Delete(savepath);
                        }

                        var writer = new VideoWriter(savepath, FourCC.DIVX, fps, frameSize);
                        
                        foreach (string f in files)
                        {
                            Mat mat = new Mat();
                            mat = Cv2.ImRead(f);
                            writer.Write(mat);
                            System.IO.File.Delete(f);
                        }
                        writer.Release();
                        System.IO.File.Delete(savepath);
                    }
                    catch
                    {
                        MessageBox.Show("静止画から動画の生成または保存に失敗しました", "ドラレコ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Environment.Exit(-1);
                    }
                    break;
                }

                // ログオフ・シャットダウンなどのとき
                var evw = new EventWatcher((o, e) =>
                {
                    timer.Dispose();
                    Environment.Exit(-1);
                });
            }


            // フラグの初期化
            isRunning = false;
            isTargetWindowExist = false;

            // メインループ
            while (true)
            {
                System.Threading.Thread.Sleep(100);

                // トップレベルウィンドウのタイトルリストを取得
                GlobalVariables.WindowList = new List<string>();
                EnumWindows(new EnumWindowsDelegate(EnumWindowCallBack), IntPtr.Zero);

                // キーワードを含むタイトルがあるか調べる
                isTargetWindowExist = false;
                foreach (string w in GlobalVariables.WindowList)
                {
                    if (w.Contains(keyword))
                    {
                        isTargetWindowExist = true;
                    }
                }

                // タイマー動作停止中/未開始で、キーワードを含むウィンドウがないときは、タイマーを開始する
                if (isRunning == false && isTargetWindowExist == false)
                {
                    timer = new System.Threading.Timer(callback, null, 0, period);
                    isRunning = true;
                }

                // タイマー動作中で、キーワードを含むウィンドウが出現したら、タイマーを停止して処理を行う
                if (isRunning == true && isTargetWindowExist == true)
                {
                    // 実行中のスレッドがすべて終了するのを待つために、確実なDisposeを行う
                    // https://so-zou.jp/software/tech/programming/c-sharp/thread/
                    // timer.Dispose();の代わりに以下のようにする
                    using (System.Threading.WaitHandle waitHandle = new System.Threading.ManualResetEvent(false))
                    {
                        if (timer.Dispose(waitHandle))
                        {
                            const int millisecondsTimeout = 1000;
                            if (!waitHandle.WaitOne(millisecondsTimeout))
                            {
                                // 指定時間内に破棄されなかった
                            }
                        }
                    }
                    isRunning = false;

                    // 動画にして保存する
                    try
                    {
                        string savepath = savedir + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".avi";

                        if (System.IO.File.Exists(savepath))
                        {
                            System.IO.File.Delete(savepath);
                        }

                        var writer = new VideoWriter(savepath, FourCC.DIVX, fps, frameSize);
                        IEnumerable<string> files = System.IO.Directory.EnumerateFiles(workdir, "*" + imageFormat);
                        foreach (string f in files)
                        {
                            Mat mat = new Mat();
                            mat = Cv2.ImRead(f);
                            writer.Write(mat);
                            System.IO.File.Delete(f);
                        }
                        writer.Release();
                    }
                    catch
                    {
                        MessageBox.Show("静止画から動画の生成または保存に失敗しました", "ドラレコ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Environment.Exit(-1);
                    }
                }

                // ログオフ・シャットダウンなどのとき
                var evw = new EventWatcher((o, e) =>
                {
                    timer.Dispose();
                    Environment.Exit(-1);
                });
            }
        }

        #region ログオフ・シャットダウンを検知する

        // https://www.exceedsystem.net/2021/02/15/how-to-handle-logoff-shutdown-reboot-events-like-winforms-formclosed-event-in-console-application/
        
        // 終了イベント監視クラス
        internal sealed class EventWatcher
        {
            public EventWatcher(FormClosedEventHandler handler)
            {
                // 別スレッドで非表示フォームを生成
                var t = new Thread((p) =>
                {
                    using (var form = new EventWatcherForm())
                    {
                        form.FormClosed += handler;
                        form.ShowDialog();
                    }
                });
                t.SetApartmentState(ApartmentState.STA);
                t.Name = "EventWatcherThread";
                t.Start();
            }

            // 非表示フォームクラス
            private sealed class EventWatcherForm : Form
            {
                // CreateParams
                protected override CreateParams CreateParams
                {
                    get
                    {
                        // 非表示フォーム用スタイル設定
                        var cp = base.CreateParams;
                        cp.Style = unchecked((int)(
                            WindowStyles.WS_POPUP
                            | WindowStyles.WS_MAXIMIZEBOX
                            | WindowStyles.WS_SYSMENU
                            | WindowStyles.WS_VISIBLE));
                        cp.ExStyle = unchecked((int)(
                            ExWindowStyles.WS_EX_TOOLWINDOW));
                        cp.Width = 0;
                        cp.Height = 0;
                        return cp;
                    }
                }

                // ウィンドウスタイル
                // https://docs.microsoft.com/en-us/windows/win32/winmsg/window-styles
                private enum WindowStyles : long
                {
                    WS_MAXIMIZEBOX = 0x00010000L,
                    WS_SYSMENU = 0x00080000L,
                    WS_VISIBLE = 0x10000000L,
                    WS_POPUP = 0x80000000L,
                }

                // 拡張ウィンドウスタイル
                // https://docs.microsoft.com/en-us/windows/win32/winmsg/extended-window-styles
                private enum ExWindowStyles : long
                {
                    WS_EX_TOOLWINDOW = 0x00000080L,
                }
            }
        }

        #endregion

        #region ウィンドウタイトルのリストを取得
            
            // https://dobon.net/vb/dotnet/process/enumwindows.html

            public class GlobalVariables
            {
                public static List<string> WindowList = new List<string>();
            }

            public delegate bool EnumWindowsDelegate(IntPtr hWnd, IntPtr lparam);

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            public extern static bool EnumWindows(EnumWindowsDelegate lpEnumFunc,
                IntPtr lparam);

            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern int GetWindowText(IntPtr hWnd,
                StringBuilder lpString, int nMaxCount);

            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern int GetWindowTextLength(IntPtr hWnd);

            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern int GetClassName(IntPtr hWnd,
                StringBuilder lpClassName, int nMaxCount);

            private static bool EnumWindowCallBack(IntPtr hWnd, IntPtr lparam)
            {
                //ウィンドウのタイトルの長さを取得する
                int textLen = GetWindowTextLength(hWnd);
                if (0 < textLen)
                {
                    //ウィンドウのタイトルを取得する
                    StringBuilder tsb = new StringBuilder(textLen + 1);
                    GetWindowText(hWnd, tsb, tsb.Capacity);

                    //ウィンドウのクラス名を取得する
                    StringBuilder csb = new StringBuilder(256);
                    GetClassName(hWnd, csb, csb.Capacity);

                    ////結果を表示する
                    //Console.WriteLine("クラス名:" + csb.ToString());
                    //Console.WriteLine("タイトル:" + tsb.ToString());

                    GlobalVariables.WindowList.Add(tsb.ToString());

                }

                //すべてのウィンドウを列挙する
                return true;
            }

        #endregion

        #region INIファイルの読み込み

            // https://smdn.jp/programming/netfx/tips/read_ini/

            private static Dictionary<string, Dictionary<string, string>> ReadIni(string file)
            {
                using (var reader = new StreamReader(file))
                {
                    var sections = new Dictionary<string, Dictionary<string, string>>(StringComparer.Ordinal);
                    var regexSection = new Regex(@"^\s*\[(?<section>[^\]]+)\].*$", RegexOptions.Singleline | RegexOptions.CultureInvariant);
                    var regexNameValue = new Regex(@"^\s*(?<name>[^=]+)=(?<value>.*?)(\s+;(?<comment>.*))?$", RegexOptions.Singleline | RegexOptions.CultureInvariant);
                    var currentSection = string.Empty;

                    // セクション名が明示されていない先頭部分のセクション名を""として扱う
                    sections[string.Empty] = new Dictionary<string, string>();

                    for (; ; )
                    {
                        var line = reader.ReadLine();

                        if (line == null)
                            break;

                        // 空行は読み飛ばす
                        if (line.Length == 0)
                            continue;

                        // コメント行は読み飛ばす
                        if (line.StartsWith(";", StringComparison.Ordinal))
                            continue;
                        else if (line.StartsWith("#", StringComparison.Ordinal))
                            continue;

                        var matchNameValue = regexNameValue.Match(line);

                        if (matchNameValue.Success)
                        {
                            // name=valueの行
                            sections[currentSection][matchNameValue.Groups["name"].Value.Trim()] = matchNameValue.Groups["value"].Value.Trim();
                            continue;
                        }

                        var matchSection = regexSection.Match(line);

                        if (matchSection.Success)
                        {
                            // [section]の行
                            currentSection = matchSection.Groups["section"].Value;

                            if (!sections.ContainsKey(currentSection))
                                sections[currentSection] = new Dictionary<string, string>();

                            continue;
                        }
                    }

                    return sections;
                }
            }

        #endregion

    }
}
