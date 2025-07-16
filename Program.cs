using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using OpenCvSharp;

namespace CourseAutomation
{
    class Program
    {
        // 系统常量
        const int SCREEN_WIDTH = 1920;
        const int SCREEN_HEIGHT = 1080;
        const int MOUSEEVENTF_LEFTDOWN = 0x02;
        const int MOUSEEVENTF_LEFTUP = 0x04;
        const int MOUSEEVENTF_WHEEL = 0x0800;
        const int WHEEL_DELTA = 120;

        // 热键常量
        const int VK_F8 = 0x77;
        const int VK_F9 = 0x78;

        // 系统坐标点
        static readonly System.Drawing.Point ENTER_TASK_POINT = new System.Drawing.Point(1712, 390);
        static readonly System.Drawing.Point SCROLL_AREA_POINT = new System.Drawing.Point(1513, 320);
        static readonly System.Drawing.Point RETURN_TASK_LIST_POINT = new System.Drawing.Point(507, 15);

        // 已完成图标相对于时长图标的偏移量
        const int COMPLETED_OFFSET_X = 544;
        const int COMPLETED_OFFSET_Y = -30;
        const int COMPLETED_CHECK_SIZE = 50;

        // 图像匹配阈值
        const double MATCH_THRESHOLD = 0.8;
        const int MIN_TEMPLATE_SIZE = 10;
        const int COURSE_DUPLICATE_THRESHOLD = 30; // 课程去重阈值

        // 程序状态
        static bool isRunning = false;
        static bool exitProgram = false;
        static int scrollCount = 0;

        [DllImport("user32.dll")]
        static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

        [DllImport("user32.dll")]
        static extern short GetAsyncKeyState(int vKey);

        [DllImport("kernel32", SetLastError = true, CharSet = CharSet.Unicode)]
        static extern IntPtr LoadLibrary(string lpFileName);

        static Program()
        {
            LoadOpenCvLibrary();
        }

        static void LoadOpenCvLibrary()
        {
            var paths = new List<string>
            {
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "runtimes", "win-x64", "native", "OpenCvSharpExtern.dll"),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OpenCvSharpExtern.dll")
            };

            foreach (var path in paths)
            {
                if (File.Exists(path))
                {
                    if (LoadLibrary(path) != IntPtr.Zero)
                    {
                        Console.WriteLine($"成功加载OpenCV库: {path}");
                        return;
                    }
                }
            }

            Console.WriteLine("警告: 无法加载OpenCvSharpExtern.dll，尝试继续运行...");
        }

        static void Main(string[] args)
        {
            Console.WriteLine("按 F8 开始程序，按 F9 退出程序");

            // 初始化测试
            try
            {
                using var testMat = new Mat();
                Console.WriteLine("OpenCV库已成功初始化");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"OpenCV初始化失败: {ex.Message}");
                Console.WriteLine("请确保已安装OpenCvSharp4和OpenCvSharp4.runtime.win NuGet包");
                Console.WriteLine("按任意键退出...");
                Console.ReadKey();
                return;
            }

            // 加载模板图像
            using var completeIcon1 = SafeLoadTemplate("1.png");
            using var completeIcon2 = SafeLoadTemplate("6.png");
            using var durationIcon = SafeLoadTemplate("2.png");
            using var passTestIcon = SafeLoadTemplate("3.png");
            using var studyCompleteIcon = SafeLoadTemplate("4.png");
            using var completedIcon = SafeLoadTemplate("5.png");

            // 任务列表位置存储
            List<System.Drawing.Point> taskPositions = new List<System.Drawing.Point>();
            int currentTaskIndex = 0;

            // 热键检测循环
            while (!exitProgram)
            {
                if (GetAsyncKeyState(VK_F8) < 0 && !isRunning)
                {
                    isRunning = true;
                    Console.WriteLine("程序开始运行...");

                    // 使用多模板识别任务列表
                    var positions1 = FindAllIcons(completeIcon1, "任务列表(模板1)");
                    var positions2 = FindAllIcons(completeIcon2, "任务列表(模板2)");

                    // 合并结果并去重排序
                    taskPositions = positions1.Concat(positions2)
                        .GroupBy(p => $"{p.X}-{p.Y}")
                        .Select(g => g.First())
                        .OrderBy(p => p.Y) // 按Y坐标排序
                        .ThenBy(p => p.X)  // 其次按X坐标排序
                        .ToList();

                    currentTaskIndex = 0;
                    Console.WriteLine($"共找到 {taskPositions.Count} 个任务项");
                }

                if (isRunning && !exitProgram)
                {
                    try
                    {
                        if (taskPositions.Count == 0 || currentTaskIndex >= taskPositions.Count)
                        {
                            Console.WriteLine("所有任务已完成，等待新任务...");
                            isRunning = false;
                            continue;
                        }

                        // 处理当前任务
                        var taskPos = taskPositions[currentTaskIndex];
                        Console.WriteLine($"\n=== 开始处理任务 {currentTaskIndex + 1}/{taskPositions.Count} ===");
                        Console.WriteLine($"任务位置: {taskPos}");

                        // 点击任务项
                        Console.WriteLine($"点击任务项位置: {taskPos}");
                        MouseClick(taskPos.X, taskPos.Y);
                        Thread.Sleep(1000);

                        // 点击进入学习页面
                        Console.WriteLine($"点击进入学习页面: ({ENTER_TASK_POINT.X}, {ENTER_TASK_POINT.Y})");
                        MouseClick(ENTER_TASK_POINT.X, ENTER_TASK_POINT.Y);
                        Thread.Sleep(2000);

                        // 处理该任务下的所有视频
                        bool taskCompleted = ProcessSingleTask(durationIcon, completedIcon,
                            passTestIcon, studyCompleteIcon, completeIcon1);

                        if (taskCompleted)
                        {
                            // 任务完成，移动到下一个任务
                            currentTaskIndex++;
                        }
                        else
                        {
                            // 任务未完成，保持当前索引
                            Console.WriteLine("任务未完成，将重试");
                        }

                        // 返回任务列表后等待
                        Thread.Sleep(1000);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"运行时错误: {ex.Message}");
                        isRunning = false;
                    }
                }

                Thread.Sleep(100);
            }

            Console.WriteLine("程序已退出");
        }

        // 安全加载模板图像
        static Mat SafeLoadTemplate(string path)
        {
            if (!File.Exists(path))
            {
                Console.WriteLine($"警告: 模板图像 {path} 不存在");
                return new Mat();
            }

            try
            {
                var template = new Mat(path, ImreadModes.Grayscale);
                if (template.Empty())
                {
                    Console.WriteLine($"模板图像 {path} 加载后为空");
                    return new Mat();
                }

                if (template.Width < MIN_TEMPLATE_SIZE || template.Height < MIN_TEMPLATE_SIZE)
                {
                    Console.WriteLine($"模板图像 {path} 尺寸过小: {template.Width}x{template.Height}");
                    return new Mat();
                }

                Console.WriteLine($"成功加载模板: {path} ({template.Width}x{template.Height})");
                return template;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"加载模板图像 {path} 失败: {ex.Message}");
                return new Mat();
            }
        }

        // 安全模板匹配
        static bool SafeMatchTemplate(Mat source, Mat template, out System.Drawing.Point? matchLocation)
        {
            matchLocation = null;

            if (source.Empty() || template.Empty())
            {
                Console.WriteLine("源图像或模板图像为空");
                return false;
            }

            if (template.Width > source.Width || template.Height > source.Height)
            {
                Console.WriteLine($"模板尺寸({template.Width}x{template.Height})大于源图像({source.Width}x{source.Height})");
                return false;
            }

            try
            {
                using var result = new Mat();
                Cv2.MatchTemplate(source, template, result, TemplateMatchModes.CCoeffNormed);

                result.MinMaxLoc(out _, out double maxVal, out _, out OpenCvSharp.Point maxLoc);

                if (maxVal > MATCH_THRESHOLD)
                {
                    matchLocation = new System.Drawing.Point(
                        maxLoc.X + template.Width / 2,
                        maxLoc.Y + template.Height / 2);
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"模板匹配失败: {ex.Message}");
            }

            return false;
        }

        // 查找所有匹配的图标
        static List<System.Drawing.Point> FindAllIcons(Mat template, string context)
        {
            var positions = new List<System.Drawing.Point>();

            if (template.Empty())
            {
                Console.WriteLine($"[{context}] 模板图像无效，跳过查找");
                return positions;
            }

            scrollCount = 0;
            int scrollAttempts = 0;

            while (scrollAttempts < 10 && !exitProgram)
            {
                using var screen = CaptureScreen();
                if (screen.Empty())
                {
                    Console.WriteLine($"[{context}] 屏幕捕获失败");
                    break;
                }

                using var screenGray = screen.CvtColor(ColorConversionCodes.BGR2GRAY);

                while (SafeMatchTemplate(screenGray, template, out var matchLoc) && matchLoc.HasValue && !exitProgram)
                {
                    positions.Add(matchLoc.Value);
                    Console.WriteLine($"找到目标位置: {matchLoc.Value}");

                    using var mask = new Mat(screenGray.Size(), MatType.CV_8UC1, Scalar.Black);
                    Cv2.Rectangle(mask,
                        new OpenCvSharp.Rect(
                            matchLoc.Value.X - template.Width / 2 - 5,
                            matchLoc.Value.Y - template.Height / 2 - 5,
                            template.Width + 10,
                            template.Height + 10),
                        Scalar.White, -1);

                    Cv2.BitwiseNot(mask, mask);
                    Cv2.BitwiseAnd(screenGray, mask, screenGray);
                }

                if (positions.Count > 0 || scrollAttempts > 5 || exitProgram)
                    break;

                MouseMove(SCROLL_AREA_POINT.X, SCROLL_AREA_POINT.Y);
                MouseScroll(-5);
                scrollCount++;
                scrollAttempts++;
                Console.WriteLine($"[{context}] 向下滚动，滚动次数: {scrollCount}");
                Thread.Sleep(1000);
            }

            Console.WriteLine($"[{context}] 找到 {positions.Count} 个目标");
            return positions;
        }

        // 处理单个任务
        // 修改后的ProcessTask方法
        static bool ProcessSingleTask(Mat durationIcon, Mat completedIcon,
               Mat passTestIcon, Mat studyCompleteIcon, Mat completeIcon)
        {
            try
            {
                Console.WriteLine($"\n--- 开始处理当前任务的所有视频 ---");

                // 移动到滚动区域
                Console.WriteLine($"移动鼠标到滚动区域坐标: ({SCROLL_AREA_POINT.X}, {SCROLL_AREA_POINT.Y})");
                MouseMove(SCROLL_AREA_POINT.X, SCROLL_AREA_POINT.Y);
                Thread.Sleep(500);

                var allVideoPositions = new List<System.Drawing.Point>();
                var processedCourses = new HashSet<int>();

                // 第0次滚动（不实际滚动）
                Console.WriteLine("开始第0次滚动检测（初始位置）");
                var initialVideos = FindVideosInView(durationIcon, completedIcon, processedCourses);
                allVideoPositions.AddRange(initialVideos);
                Console.WriteLine($"初始位置找到 {initialVideos.Count} 个未完成视频");

                // 第一次实际滚动
                Console.WriteLine("执行第1次滚动");
                MouseScroll(-10);
                Thread.Sleep(1500);

                // 第一次滚动后检测
                var scrolledVideos = FindVideosInView(durationIcon, completedIcon, processedCourses);
                allVideoPositions.AddRange(scrolledVideos);
                Console.WriteLine($"滚动后找到 {scrolledVideos.Count} 个新增未完成视频");
                MouseScroll(10);
                Thread.Sleep(1500);

                // 合并并去重（基于Y坐标）
                var finalVideoPositions = allVideoPositions
                    .GroupBy(p => p.Y / COURSE_DUPLICATE_THRESHOLD) // 按Y坐标分组去重
                    .Select(g => g.First())
                    .OrderBy(p => p.Y)
                    .ThenBy(p => p.X)
                    .ToList();

                Console.WriteLine($"\n合并后共发现 {finalVideoPositions.Count} 个未完成视频（已去重）");

                // 处理所有视频
                for (int i = 0; i < finalVideoPositions.Count; i++)
                {
                    Console.WriteLine($"\n处理视频 {i + 1}/{finalVideoPositions.Count}");
                    Console.WriteLine($"点击视频位置: {finalVideoPositions[i]}");
                    MouseClick(finalVideoPositions[i].X, finalVideoPositions[i].Y);
                    Thread.Sleep(3000);

                    bool videoCompleted = false;
                    while (!videoCompleted)
                    {
                        using var screen = CaptureScreen();
                        if (screen.Empty()) continue;

                        using var screenGray = screen.CvtColor(ColorConversionCodes.BGR2GRAY);

                        // 检测"点我通过测试"
                        if (SafeMatchTemplate(screenGray, passTestIcon, out var passTestLoc) && passTestLoc.HasValue)
                        {
                            Console.WriteLine($"检测到[通过测试]，点击位置: {passTestLoc.Value}");
                            MouseClick(passTestLoc.Value.X, passTestLoc.Value.Y);
                            Thread.Sleep(1000);
                        }

                        // 检测"学习完成"
                        if (SafeMatchTemplate(screenGray, studyCompleteIcon, out var completeLoc) && completeLoc.HasValue)
                        {
                            videoCompleted = true;
                            Console.WriteLine("视频学习完成");
                            Thread.Sleep(2000);
                        }
                        else
                        {
                            Thread.Sleep(1000);
                        }
                    }
                }

                Console.WriteLine("\n当前任务所有视频处理完成，返回任务列表");
                MouseClick(RETURN_TASK_LIST_POINT.X, RETURN_TASK_LIST_POINT.Y);
                Thread.Sleep(2000);

                // 验证是否成功返回
                using var verificationScreen = CaptureScreen();
                if (!verificationScreen.Empty())
                {
                    using var grayScreen = verificationScreen.CvtColor(ColorConversionCodes.BGR2GRAY);
                    if (SafeMatchTemplate(grayScreen, completeIcon, out _))
                    {
                        Console.WriteLine("成功返回任务列表");
                        return true;
                    }
                }

                Console.WriteLine("返回任务列表失败");
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"任务处理异常: {ex.Message}");
                return false;
            }
        }

        // 新增方法：在当前视图中查找未完成视频
        static List<System.Drawing.Point> FindVideosInView(Mat durationIcon, Mat completedIcon, HashSet<int> processedCourses)
        {
            var videoPositions = new List<System.Drawing.Point>();

            using var screen = CaptureScreen();
            if (screen.Empty()) return videoPositions;

            using var screenGray = screen.CvtColor(ColorConversionCodes.BGR2GRAY);

            // 查找所有时长图标
            while (SafeMatchTemplate(screenGray, durationIcon, out var durationLoc) &&
                   durationLoc.HasValue)
            {
                int currentY = durationLoc.Value.Y;

                // 去重检查
                if (processedCourses.Any(y => Math.Abs(y - currentY) < COURSE_DUPLICATE_THRESHOLD))
                {
                    Console.WriteLine($"  → 已处理过该课程(Y={currentY})，跳过");

                    // 屏蔽已找到的区域
                    using var mask = new Mat(screenGray.Size(), MatType.CV_8UC1, Scalar.Black);
                    Cv2.Rectangle(mask,
                        new OpenCvSharp.Rect(
                            durationLoc.Value.X - durationIcon.Width / 2 - 5,
                            currentY - durationIcon.Height / 2 - 5,
                            durationIcon.Width + 10,
                            durationIcon.Height + 10),
                        Scalar.White, -1);

                    Cv2.BitwiseNot(mask, mask);
                    Cv2.BitwiseAnd(screenGray, mask, screenGray);
                    continue;
                }

                // 记录已处理的课程位置
                processedCourses.Add(currentY);

                // 立即检查该课程是否已完成
                if (!CheckCompletedAtPosition(screen, durationLoc.Value, completedIcon))
                {
                    videoPositions.Add(durationLoc.Value);
                    Console.WriteLine($"  → 发现未完成视频 @ {durationLoc.Value}");
                }

                // 屏蔽已找到的区域
                using var mask2 = new Mat(screenGray.Size(), MatType.CV_8UC1, Scalar.Black);
                Cv2.Rectangle(mask2,
                    new OpenCvSharp.Rect(
                        durationLoc.Value.X - durationIcon.Width / 2 - 5,
                        durationLoc.Value.Y - durationIcon.Height / 2 - 5,
                        durationIcon.Width + 10,
                        durationIcon.Height + 10),
                    Scalar.White, -1);

                Cv2.BitwiseNot(mask2, mask2);
                Cv2.BitwiseAnd(screenGray, mask2, screenGray);
            }

            return videoPositions;
        }

        // 检查课程是否已完成（在时长图标位置附近检测）
        static bool CheckCompletedAtPosition(Mat screen, System.Drawing.Point durationPoint, Mat completedIcon)
        {
            if (completedIcon.Empty())
            {
                Console.WriteLine("已完成图标模板无效");
                return false;
            }

            // 计算已完成图标的检查位置（时长图标右544px，上30px）
            int checkX = durationPoint.X + COMPLETED_OFFSET_X;
            int checkY = durationPoint.Y + COMPLETED_OFFSET_Y;

            Console.WriteLine($"检查完成状态: 时长位置({durationPoint.X},{durationPoint.Y}) → 检查位置({checkX},{checkY})");

            // 调整ROI区域大小，确保不小于模板尺寸
            int roiWidth = Math.Max(completedIcon.Width + 20, COMPLETED_CHECK_SIZE);
            int roiHeight = Math.Max(completedIcon.Height + 20, COMPLETED_CHECK_SIZE);

            int roiX = Math.Max(0, checkX - roiWidth / 2);
            int roiY = Math.Max(0, checkY - roiHeight / 2);
            roiWidth = Math.Min(roiWidth, screen.Width - roiX);
            roiHeight = Math.Min(roiHeight, screen.Height - roiY);

            if (roiWidth <= 0 || roiHeight <= 0)
            {
                Console.WriteLine($"无效的ROI区域: {roiX},{roiY},{roiWidth},{roiHeight}");
                return false;
            }

            Console.WriteLine($"ROI区域: {roiWidth}x{roiHeight} (模板: {completedIcon.Width}x{completedIcon.Height})");

            using var roi = new Mat(screen, new OpenCvSharp.Rect(roiX, roiY, roiWidth, roiHeight));
            using var roiGray = roi.CvtColor(ColorConversionCodes.BGR2GRAY);

            // 检查是否匹配"已完成"图标
            bool isCompleted = SafeMatchTemplate(roiGray, completedIcon, out _);
            Console.WriteLine($"  → 完成状态: {(isCompleted ? "已完成" : "未完成")}");
            return isCompleted;
        }

        // 屏幕捕获
        static Mat CaptureScreen()
        {
            try
            {
                using var bitmap = new Bitmap(SCREEN_WIDTH, SCREEN_HEIGHT);
                using (var g = Graphics.FromImage(bitmap))
                {
                    g.CopyFromScreen(0, 0, 0, 0, new System.Drawing.Size(SCREEN_WIDTH, SCREEN_HEIGHT));
                }

                var rect = new Rectangle(0, 0, bitmap.Width, bitmap.Height);
                BitmapData bmpData = bitmap.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
                try
                {
                    Mat mat = Mat.FromPixelData(
                        bmpData.Height,
                        bmpData.Width,
                        MatType.CV_8UC4,
                        bmpData.Scan0,
                        bmpData.Stride);

                    Mat bgrMat = new Mat();
                    Cv2.CvtColor(mat, bgrMat, ColorConversionCodes.BGRA2BGR);
                    return bgrMat;
                }
                finally
                {
                    bitmap.UnlockBits(bmpData);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"屏幕捕获失败: {ex.Message}");
                return new Mat();
            }
        }

        // 鼠标操作
        static void MouseMove(int x, int y) => SetCursorPos(x, y);

        static void MouseClick(int x, int y)
        {
            MouseMove(x, y);
            mouse_event(MOUSEEVENTF_LEFTDOWN, x, y, 0, 0);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_LEFTUP, x, y, 0, 0);
            Thread.Sleep(500);
        }

        static void MouseScroll(int steps)
        {
            int scrollAmount = steps * WHEEL_DELTA;
            mouse_event(MOUSEEVENTF_WHEEL, 0, 0, scrollAmount, 0);
            Thread.Sleep(300);
        }
    }
}