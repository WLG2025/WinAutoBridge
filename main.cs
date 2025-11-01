namespace p5
{
    using System.Net; // HttpListener
    using System.Text.Json; // JsonSerializer
    using System.Text; // Encoding
    using System.Diagnostics; // Process
    using System.Runtime.InteropServices; // DllImport


    // 简单的日志记录类
    public class Logger
    {
        private readonly string _logFilePath;
        private readonly string _fileNameBase = string.Empty;
        private string _lastDate = string.Empty;
        private readonly object _lockObject = new();
        private readonly int _processId;
        public Logger(string fileName)
        {
            string logDirectory = Path.Combine(
                "./",
                "log"
            );
            Directory.CreateDirectory(logDirectory);

            _logFilePath = logDirectory;
            _fileNameBase = fileName.Split(".log")[0];
            _lastDate = DateTime.Now.ToString("yyyyMMdd");
            _processId = Environment.ProcessId;
        }

        public void Debug(string message)
        {
            WriteLog("DEBUG", message);
        }

        public void Info(string message)
        {
            WriteLog("INFO", message);
        }

        public void Warn(string message)
        {
            WriteLog("WARN", message);
        }

        public void Error(string message)
        {
            WriteLog("ERROR", message);
        }

        private void WriteLog(string level, string message)
        {
            try
            {
                lock (_lockObject)
                {
                    string nowDate = DateTime.Now.ToString("yyyyMMdd");
                    if (nowDate != _lastDate)
                    {
                        _lastDate = nowDate;
                    }
                    string logFilePath = Path.Combine(_logFilePath, $"{_fileNameBase}_{_lastDate}.log");
                    string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}|{_processId}|{level}|{message}{Environment.NewLine}";
                    File.AppendAllText(logFilePath, logEntry);
                }
            }
            catch
            {
                // 忽略日志记录错误
            }
        }
    }

    class Program
    {
        private static readonly Logger logger = new("admin");
        private static bool isFinish = true;
        private static TextBox? processNameTextBox;
        private static TextBox? windowTitleTextBox;
        private static TextBox? msgContentTextBox;


        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
        // 委托声明
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
        // 获取指定进程的所有窗口句柄
        public static List<IntPtr> GetAllWindowsFromProcess(int processId)
        {
            List<IntPtr> windowHandles = [];
            EnumWindows((hWnd, param) =>
            {
                _ = GetWindowThreadProcessId(hWnd, out int pid);
                if (pid == processId)
                {
                    windowHandles.Add(hWnd);
                }
                return true;
            }, IntPtr.Zero);
            return windowHandles;
        }


        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowTextLength(IntPtr hWnd);
        private static string GetWindowTextEx(IntPtr hWnd)
        {
            int length = GetWindowTextLength(hWnd);
            if (length == 0)
            {
                return string.Empty;
            }

            StringBuilder sb = new(length + 1);
            _ = GetWindowText(hWnd, sb, sb.Capacity);
            return sb.ToString();
        }


        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool SetFocus(IntPtr hWnd);
        // Windows常量
        private const int SW_RESTORE = 9;
        private const uint WM_ACTIVATE = 0x0006;
        private const uint WA_ACTIVE = 1;


        [DllImport("user32.dll")]
        private static extern bool SetCursorPos(int X, int Y);
        [DllImport("user32.dll")]
        private static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);
        // 鼠标事件常量
        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;
        public static void ClickAtPosition(int x, int y)
        {
            // 设置鼠标位置
            SetCursorPos(x, y);

            // 模拟鼠标左键按下和释放
            mouse_event(MOUSEEVENTF_LEFTDOWN, x, y, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, x, y, 0, 0);
        }


        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
        public static Point GetWindowBottomRight(IntPtr hWnd)
        {
            if (GetWindowRect(hWnd, out RECT rect))
            {
                return new Point(rect.Right, rect.Bottom);
            }
            else
            {
                // 如果获取失败，返回默认坐标(0,0)
                return new Point(0, 0);
            }
        }


        // 添加HTTP监听和JSON处理功能
        private static HttpListener? httpListener;
        private static bool isListening = false;
        // 启动HTTP服务器处理POST JSON数据
        public static void StartHttpServer(int port = 8080)
        {
            try
            {
                httpListener = new HttpListener();
                httpListener.Prefixes.Add($"http://localhost:{port}/");
                httpListener.Prefixes.Add($"http://127.0.0.1:{port}/");
                httpListener.Start();
                isListening = true;

                logger.Info($"HTTP服务器已启动|监听端口:{port}");

                // 异步处理请求
                Task.Run(async () =>
                {
                    while (isListening)
                    {
                        try
                        {
                            var context = await httpListener.GetContextAsync();
                            _ = Task.Run(() => ProcessRequest(context));
                        }
                        catch (HttpListenerException)
                        {
                            // 服务器已停止时会抛出异常，正常退出循环
                            break;
                        }
                        catch (ObjectDisposedException)
                        {
                            break;
                        }
                        catch (Exception ex)
                        {
                            logger.Error($"处理HTTP请求时发生错误: {ex.Message}|{ex.GetType().FullName}");
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                logger.Error($"启动HTTP服务器失败: {ex.Message}");
            }
        }
        // 在Program类外部或内部添加响应类定义
        public class JsonResponse
        {
            public string Status { get; set; } = string.Empty;
            public string Message { get; set; } = string.Empty;
            public string? Data { get; set; }
            public string Time { get; set; } = string.Empty;
        }
        // 处理HTTP请求
        private static async void ProcessRequest(HttpListenerContext context)
        {
            var request = context.Request;
            var response = context.Response;
            try
            {
                // 只处理POST请求
                if (request.HttpMethod != "POST")
                {
                    response.StatusCode = 405; // Method Not Allowed
                    response.Close();
                    return;
                }

                // 检查内容类型是否为JSON
                if ((request.ContentType == null) || !request.ContentType.Contains("application/json"))
                {
                    response.StatusCode = 415; // Unsupported Media Type
                    response.Close();
                    return;
                }

                // 读取JSON数据
                using var reader = new StreamReader(request.InputStream, request.ContentEncoding);
                string jsonBody = await reader.ReadToEndAsync();

                logger.Info($"收到POST JSON请求: {jsonBody}|path:{request.Url?.PathAndQuery}");

                // 将JSON数据发送到剪贴板
                SendToClipboard(jsonBody, "post");
                _ = Task.Run(() =>
                {
                    try
                    {
                        Thread.Sleep(2000);
                        string processName = (processNameTextBox != null) ? processNameTextBox.Text : "";
                        string windowTitle = (windowTitleTextBox != null) ? windowTitleTextBox.Text : "";
                        SendCustomMessage(processName, windowTitle, false);
                    }
                    catch (Exception ex)
                    {
                        logger.Error($"使用剪贴板转发消息异常|{ex.Message}");
                    }
                });

                // 处理JSON数据并创建标准JSON响应
                var responseObj = new JsonResponse
                {
                    Status = "success",
                    Message = "OK",
                    Data = jsonBody,
                    Time = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                };

                string responseString = JsonSerializer.Serialize(responseObj);

                // 返回响应
                byte[] buffer = Encoding.UTF8.GetBytes(responseString);
                response.ContentType = "application/json";
                response.ContentLength64 = buffer.Length;
                response.StatusCode = 200;

                await response.OutputStream.WriteAsync(buffer);
                response.OutputStream.Close();
            }
            catch (Exception ex)
            {
                logger.Error($"处理请求时发生错误: {ex.Message}");
                response.StatusCode = 500;
                response.Close();
            }
            finally
            {
                response.Close();
            }
        }
        // 停止HTTP服务器
        public static void StopHttpServer()
        {
            isListening = false;
            try
            {
                if (httpListener != null)
                {
                    if (httpListener.IsListening)
                    {
                        httpListener.Stop();
                    }
                    httpListener.Close();
                    // if (httpListener is IDisposable disposable)
                    // {
                    //     disposable.Dispose();
                    // }
                    httpListener = null;
                }
            }
            catch (Exception ex)
            {
                logger.Error($"停止HTTP服务器时发生错误|{ex}");
            }

            logger.Info("HTTP服务器已停止");
        }

        public static void SendToClipboard(string text, string sFrom)
        {
            try
            {
                // 需要在STA线程中操作剪贴板，所以使用Invoke
                var clipboardThread = new Thread(() =>
                {
                    try
                    {
                        Clipboard.SetText(text);
                        logger.Info($"已发送消息内容到剪贴板|len:{text.Length}|{sFrom}");
                    }
                    catch (Exception clipboardEx)
                    {
                        logger.Error($"写入剪贴板失败|{clipboardEx.Message}|{sFrom}");
                    }
                });
                clipboardThread.SetApartmentState(ApartmentState.STA);
                clipboardThread.Start();
                clipboardThread.Join(); // 等待操作完成
            }
            catch (Exception ex)
            {
                logger.Error($"处理剪贴板操作时发生错误|{ex.Message}|{sFrom}");
            }
        }

        public struct WindowHandleInfo(IntPtr handle, int processId)
        {
            public IntPtr Handle { get; set; } = handle;
            public int ProcessId { get; set; } = processId;
        }
        public static WindowHandleInfo FindWindowHandle(string processName, string windowTitle)
        {
            try
            {
                Process[] processList = Process.GetProcessesByName(processName);
                if (processList.Length <= 0)
                {
                    logger.Info($"未找到进程: {processName}");
                    return new WindowHandleInfo(IntPtr.Zero, 0);
                }

                foreach (Process process in processList)
                {
                    IntPtr dstHandle = IntPtr.Zero;
                    List<IntPtr> windowHandles = GetAllWindowsFromProcess(process.Id);
                    foreach (IntPtr hwnd in windowHandles)
                    {
                        string title = GetWindowTextEx(hwnd);
                        if (title == windowTitle)
                        {
                            dstHandle = hwnd;
                            break;
                        }
                    }
                    if (dstHandle != IntPtr.Zero)
                    {
                        logger.Info($"找到窗口: {windowTitle}|handle:{dstHandle}|pid:{process.Id}");
                        return new WindowHandleInfo(dstHandle, process.Id);
                    }
                }
                logger.Info($"未找到窗口: {windowTitle}|{processName}|process count: {processList.Length}");
            }
            catch (Exception ex)
            {
                logger.Error($"查找窗口错误: {ex.Message}|{processName}|{windowTitle}");
            }
            return new WindowHandleInfo(IntPtr.Zero, 0);
        }

        public static bool ActivateWindow(WindowHandleInfo info)
        {
            if (info.Handle == IntPtr.Zero)
            {
                return false;
            }

            try
            {
                // 恢复窗口（如果最小化）
                ShowWindow(info.Handle, SW_RESTORE);

                // 将窗口置于前台
                SetForegroundWindow(info.Handle);

                // 激活窗口
                SendMessage(info.Handle, WM_ACTIVATE, new IntPtr(WA_ACTIVE), IntPtr.Zero);

                // 设置焦点到窗口
                SetFocus(info.Handle);

                // 等待窗口完全激活
                Thread.Sleep(1000);
                logger.Info($"成功激活窗口|handle:{info.Handle}|pid:{info.ProcessId}");
                return true;
            }
            catch (Exception ex)
            {
                logger.Error($"激活窗口失败: {ex.Message}|handle:{info.Handle}|pid:{info.ProcessId}");
            }
            return false;
        }

        public static void DoSendWxMessage(WindowHandleInfo info)
        {
            if (info.Handle == IntPtr.Zero)
            {
                return;
            }

            try
            {
                // 发送命令前，确保在编辑区单击左键
                // ClickAtPosition(1200, 500); // 这句很关键

                // 发送 Ctrl+V 快捷键执行粘贴操作
                SendKeys.SendWait("^v");

                {
                    // 微信窗口右下角到发送按钮的偏移量
                    int xOffset = 62;
                    int yOffset = 31;
                    Point bottomRight = GetWindowBottomRight(info.Handle);
                    ClickAtPosition(bottomRight.X - xOffset, bottomRight.Y - yOffset);
                }
                logger.Info("已发送Ctrl+V快捷键");
            }
            catch (Exception ex)
            {
                logger.Error($"发送Ctrl+V快捷键失败: {ex.Message}");
            }
        }

        public static void SendCustomMessage(string processName, string windowTitle, bool bCheck = true)
        {
            if ((processName == "") || (windowTitle == ""))
            {
                logger.Info($"进程名或窗口标题不能为空|{processName}|{windowTitle}");
                return;
            }

            if (bCheck)
            {
                // 检查剪贴板内容
                if (Clipboard.ContainsText())
                {
                    string clipboardText = Clipboard.GetText();
                    logger.Info($"剪贴板内容长度: {clipboardText.Length}|内容: {clipboardText}");
                }
                else
                {
                    logger.Warn("剪贴板中没有文本内容");
                    return;
                }
            }

            var info = FindWindowHandle(processName, windowTitle);
            if (ActivateWindow(info))
            {
                DoSendWxMessage(info);
            }
        }


        public static Form InitApp()
        {
            string appDirectory = AppDomain.CurrentDomain.BaseDirectory;
            // 切换到程序所在目录
            Directory.SetCurrentDirectory(appDirectory);
            logger.Info($"应用程序启动|{appDirectory}");

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // 启动HTTP服务器
            StartHttpServer(58080);

            return InitUI();
        }

        public static Form InitUI()
        {
            // 创建主窗口
            Form form = new()
            {
                Text = "Admin Tool",
                Size = new Size(500, 200)
            };

            AddTablePanel(form);

            return form;
        }

        public static void AddTablePanel(Form form)
        {
            // 创建TableLayoutPanel
            TableLayoutPanel tableLayoutPanel = new()
            {
                Dock = DockStyle.None,
                Location = new Point(10, 10),
                Size = new Size(form.ClientSize.Width - 20, form.ClientSize.Height - 20),
                ColumnCount = 2,
                RowCount = 4,
                // Margin = new Padding(10, 10, 0, 0),
                // Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top
            };

            // 设置列宽和行高
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80)); // 标签列
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); // 输入框列
            // tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 10));
            // tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 10));
            // tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 10));
            // tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 10));

            // 添加控件
            tableLayoutPanel.Controls.Add(new Label
            {
                Text = "进程名:",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            }, 0, 0);
            processNameTextBox = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Text = "weixin" };
            tableLayoutPanel.Controls.Add(processNameTextBox, 1, 0);

            tableLayoutPanel.Controls.Add(new Label
            {
                Text = "窗口名:",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            }, 0, 1);
            windowTitleTextBox = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Text = "空无一" };
            tableLayoutPanel.Controls.Add(windowTitleTextBox, 1, 1);

            tableLayoutPanel.Controls.Add(new Label
            {
                Text = "消息内容:",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            }, 0, 2);
            msgContentTextBox = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right };
            tableLayoutPanel.Controls.Add(msgContentTextBox, 1, 2);

            Button btn = new()
            {
                Text = "测试发送",
                Anchor = AnchorStyles.Right,
                Size = new Size(100, 35),
            };
            btn.Click += (sender, e) =>
            {
                if (!isFinish)
                {
                    return;
                }
                isFinish = false;
                logger.Info("测试发送消息");
                CopyMessage();
                SendCustomMessage(processNameTextBox.Text, windowTitleTextBox.Text);
                isFinish = true;
            };
            tableLayoutPanel.Controls.Add(btn, 1, 3);

            form.Controls.Add(tableLayoutPanel);
        }

        public static void CopyMessage()
        {
            if (msgContentTextBox == null)
            {
                return;
            }
            string text = msgContentTextBox.Text;
            if (text == "")
            {
                return;
            }
            SendToClipboard(text, "copy");
        }

        public static void DeInitApp()
        {
            // 停止HTTP服务器
            StopHttpServer();

            logger.Info("应用程序关闭");
        }

        [STAThread]
        static void Main(string[] args)
        {
            Application.Run(InitApp());
            DeInitApp();
        }
    }
}
