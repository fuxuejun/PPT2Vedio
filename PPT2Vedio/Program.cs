using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Configuration;

namespace PPT2Vedio
{
    static class Program
    {
        // 处理文件夹临时路径
        private static String tempPath = ConfigurationManager.AppSettings["tempPath"];
        // 处理结果保存路径
        private static String resultPath = ConfigurationManager.AppSettings["resultPath"];

        static void Main(string[] args)
        {
            // 监测文件夹路径
            var moniterPath = ConfigurationManager.AppSettings["moniterPath"];

            // 初始化文件夹监控线程
            var folderMoniterThread = InitFolderMoniterThread();
            folderMoniterThread.Start(moniterPath);
        }

        /// <summary>
        /// 文件夹监控线程
        /// </summary>
        private static Thread InitFolderMoniterThread()
        {
            var thread = new Thread(path =>
                {
                    while (true)
                    {
                        // 获取监控路径中的文件数量
                        String[] files = Directory.GetFiles(path.ToString(), "*.ppt*");

                        var builder = new StringBuilder();

                        foreach (var file in files)
                        {
                            builder.Append(file + "  ");
                        }

                        LogFactory.Info("获取文件数量:{0}  {1}", files.Length, builder);

                        // 先把文件转移
                        foreach (var file in files)
                        {
                            var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(file);

                            // 如果路径有中文，替换为guid
                            if (HasChinese(fileNameWithoutExtension))
                            {
                                fileNameWithoutExtension = Guid.NewGuid().ToString().Replace("-", "");
                            }

                            var path1 = Path.Combine(tempPath, fileNameWithoutExtension);

                            if (!Directory.Exists(path1))
                            {
                                Directory.CreateDirectory(path1);
                            }

                            LogFactory.Info("移动文件到新路径:{0}", path1);

                            var newPath = Path.Combine(path1, Path.GetFileName(file));

                            // 复制文件
                            File.Move(file, newPath);

                            LogFactory.Info("开始处理文件:{0}", file);
                            OperateFile(newPath);
                        }

                        // 3秒检查一次
                        Thread.Sleep(1000 * 3);
                    }
                });

            return thread;
        }

        /// <summary>
        /// 初始化处理文件线程
        /// </summary>
        /// <returns></returns>
        private static void OperateFile(String path)
        {
            var rootPath = Path.GetDirectoryName(path);

            // 启动转换Image程序
            CallProcess("ppt2image.exe", String.Format("-f {0} -o {1}", path, rootPath), false);

            // 杀死进程
            KillProcess("ppt2image");

            // 生成结果文件
            GernateResult(path);

            // 删除临时文件
            Directory.Delete(rootPath, true);
        }

        /// <summary>
        /// 生成结果文件
        /// </summary>
        /// <param name="path"></param>
        private static void GernateResult(String path)
        {
            var rootPath = Path.GetDirectoryName(path);

            var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(path);

            if (!Directory.Exists(Path.Combine(resultPath, fileNameWithoutExtension)))
            {
                Directory.CreateDirectory(Path.Combine(resultPath, fileNameWithoutExtension));
            }

            // 复制备注结果文件
            File.Copy(Path.Combine(rootPath, fileNameWithoutExtension + "_note.txt"),
                      Path.Combine(resultPath, fileNameWithoutExtension, "note.txt"), true);


            // 复制文件
            File.Copy("ffmpeg.exe", Path.Combine(rootPath, "ffmpeg.exe"), true);


            var resultFilePath = Path.Combine(rootPath, fileNameWithoutExtension + "_result.txt");

            var list = GetFileToList(resultFilePath);

            var files = Directory.GetFiles(rootPath, "*.jpg");
            
            for (int i = 0; i < list.Count - 1; i++)
            {
                var str1 = list[i];
                var str2 = list[i + 1];

                var filePrefix = str1.Split(';')[1].Split('@')[0];
                var startIndex = Int32.Parse(str1.Split(';')[0]) + 1;
                var endIndex = Int32.Parse(str2.Split(';')[0]);
                try
                {
                    ChangeFileName(startIndex, endIndex, files, filePrefix);
                }
                catch (Exception ex)
                {
                    LogFactory.Error("文件改名异常:" + ex.Message);
                }
            }
            
            files = Directory.GetFiles(rootPath, "*.jpg");

            var startList = new List<String>();
            var tempIndex = 1;
            var tempPrefix = "";

            foreach (var str in list)
            {
                var t = str.Split(';')[1].Split('@')[0];

                var tt = t.Split('-')[0];

                if (!startList.Contains(tt))
                {
                    startList.Add(tt);
                }

                if (tempPrefix != tt)
                {
                    tempPrefix = tt;
                    tempIndex = 1;
                }

                foreach (var file in files)
                {
                    if (Path.GetFileNameWithoutExtension(file).StartsWith(t))
                    {
                        try
                        {
                            File.Move(file, Path.Combine(Path.GetDirectoryName(file), tt + "-" + tempIndex.ToString("0000") + ".jpg"));
                            tempIndex++;
                        }
                        catch (Exception ex)
                        {
                            // LogFactory.Error("文件复制异常:" + ex.Message);
                        }
                    }
                }
            }

            try
            {
                foreach (var str in startList)
                {
                    // mp4临时文件
                    var mp4TempFilePath = Path.Combine(rootPath, fileNameWithoutExtension + "-" + str + ".mp4");

                    // 启动转换Vedio程序，转换每个页面为视频
                    CallProcess(Path.Combine(rootPath, "ffmpeg.exe"),
                                                String.Format("-i \"{0}\" -r 25 -f mp4 -s 1024x768 -b 300k -vcodec h264 {1}",
                                                              Path.Combine(rootPath, str + "-%04d.jpg"),
                                                              mp4TempFilePath), false);

                    File.Copy(mp4TempFilePath, Path.Combine(resultPath, fileNameWithoutExtension, str + ".mp4"), true);
                }
            }
            catch (Exception ex)
            {
                LogFactory.Error("生成MP4文件异常:" + ex.Message);
            }

            // 复制图片文件
            files = Directory.GetFiles(Path.Combine(rootPath, "images"));

            for (int index = 0; index < files.Length; index++)
            {
                try
                {
                    var file = files[index];
                    File.Move(file, Path.Combine(resultPath, fileNameWithoutExtension, (index + 1) + ".png"));
                }
                catch (Exception ex)
                {
                    LogFactory.Error("复制PPT图片报错:" + ex.Message);
                }

            }
            try
            {
                SaveResultFile(resultFilePath, list);

                File.Copy(resultFilePath,
                    Path.Combine(resultPath, fileNameWithoutExtension, "result.txt"), true);
            }
            catch (Exception ex)
            {
                LogFactory.Error("复制结果文件报错:" + ex.Message);
            }
        }

        /// <summary>
        /// 改变文件名称
        /// </summary>
        /// <param name="startIndex"></param>
        /// <param name="endIndex"></param>
        /// <param name="files"></param>
        /// <param name="filePrefix"></param>
        private static void ChangeFileName(Int32 startIndex, Int32 endIndex, String[] files, String filePrefix)
        {
            var tempEndIndex = startIndex + 94;

            Int32 index = 1;
            foreach (var file in files)
            {
                var fileIndex = Int32.Parse(Path.GetFileNameWithoutExtension(file));

                if (fileIndex >= startIndex &&
                    fileIndex <= tempEndIndex)
                {
                    if (fileIndex <= endIndex)
                    {
                        // 如果在最后的索引范围内，直接复制
                        File.Copy(file, Path.Combine(Path.GetDirectoryName(file), filePrefix + "-" + index.ToString("000") + ".jpg"), true);
                    }
                    else
                    {
                        // 如果超出索引，复制最后一个文件
                        File.Copy(files[endIndex - 1], Path.Combine(Path.GetDirectoryName(file), filePrefix + "-" + index.ToString("000") + ".jpg"), true);
                    }


                    index++;
                }
                else
                {
                    index = 1;
                }
            }
        }

        /// <summary>
        /// 读取文件到集合中
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        private static List<String> GetFileToList(String filePath)
        {
            var list = new List<String>();

            // 读取结果文件
            var stream = new FileStream(filePath, FileMode.Open);
            var reader = new StreamReader(stream);

            while (!reader.EndOfStream)
            {
                list.Add(reader.ReadLine());
            }

            reader.Close();
            stream.Close();

            reader.Dispose();
            stream.Dispose();

            return list;
        }

        /// <summary>
        /// 保存结果文件
        /// </summary>
        private static void SaveResultFile(String resultFilePath, IEnumerable<String> result)
        {
            var stream = new FileStream(resultFilePath, FileMode.Create);
            var writer = new StreamWriter(stream);

            var temp = new StringBuilder("{");

            foreach (var str in result)
            {
                //writer.WriteLine(str.Split(';')[1]);
                var t = str.Split(';')[1].Split('-');
                var page = t[0];
                var animation = t[1].Split('@')[0];
                var duration = t[1].Split('@')[1];

                temp.Append("{\"page\":\"" + page + "\",\"animation\":\"" +
                            animation + "\"，\"duration\":\"" + duration + "\"},");

                writer.Flush();
                stream.Flush();
            }

            temp = temp.Remove(temp.Length - 1, 1).Append("}");

            writer.Write(temp);

            writer.Close();
            stream.Close();

            writer.Dispose();
            stream.Dispose();
        }

        /// <summary>
        /// 保存结果文件
        /// </summary>
        private static void SaveResultFile(String resultFilePath, String resultStr)
        {
            var reg = new Regex(@"time=(?<time>\d{2}:\d{2}:\d{2}.\d{2})");
            var matches = reg.Matches(resultStr);

            var stream = new FileStream(resultFilePath, FileMode.Create);
            var writer = new StreamWriter(stream);

            for (int i = 0; i < matches.Count; i++)
            {
                Match match = matches[i];
                writer.WriteLine(i + ":" + match.Groups["time"].Value);
            }

            writer.Close();
            stream.Close();

            writer.Dispose();
            stream.Dispose();
        }

        /// <summary>
        /// 判断是否有中文
        /// </summary>
        /// <param name="words"></param>
        /// <returns></returns>
        public static bool HasChinese(string words)
        {
            string temp;
            for (int i = 0; i < words.Length; i++)
            {
                temp = words.Substring(i, 1);
                byte[] sarr = System.Text.Encoding.GetEncoding("gb2312").GetBytes(temp);
                if (sarr.Length == 2)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 杀死进程
        /// </summary>
        /// <param name="processName"></param>
        private static void KillProcess(String processName)
        {
            foreach (var process in Process.GetProcesses())
            {
                if (process.ProcessName.ToLower().Contains(processName))
                {
                    process.Kill();
                }
            }
        }

        /// <summary>
        /// 启动程序
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="args"></param>
        /// <param name="isGetResults"> </param>
        private static String CallProcess(string fileName, string args, Boolean isGetResults)
        {
            var process = new Process();
            process.StartInfo.FileName = fileName;//设置运行的命令行文件
            process.StartInfo.Arguments = args;//设置命令参数
            process.StartInfo.CreateNoWindow = true;//不显示dos命令行窗口
            process.StartInfo.UseShellExecute = false;//是否指定操作系统外壳进程启动程序
            process.StartInfo.RedirectStandardOutput = isGetResults;
            process.StartInfo.RedirectStandardError = isGetResults;

            // 启动
            process.Start();

            var result = new StringBuilder();

            if (isGetResults)
            {
                process.OutputDataReceived += (s, e) => result.AppendLine(e.Data);
                process.ErrorDataReceived += (s, e) => result.AppendLine(e.Data);

                process.BeginErrorReadLine();
                process.BeginOutputReadLine();
            }


            // 等待完成
            process.WaitForExit();

            if (isGetResults)
            {
                process.CancelOutputRead();
                process.CancelErrorRead();

                return result.ToString();
            }
            else
            {
                return "";
            }
        }
    }
}
