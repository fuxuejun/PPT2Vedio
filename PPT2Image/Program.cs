using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Diagnostics;
using ThreadState = System.Threading.ThreadState;
using System.IO;

namespace PPT2Image
{
    static class Program
    {
        [DllImport("Gdi32.dll")]
        private static extern int BitBlt(IntPtr hDC, int x, int y, int nWidth, int nHeight, IntPtr hSrcDC, int xSrc, int ySrc, int dwRop);
        
        private static Object obj = new object();
        /// <summary>
        /// 图片集合
        /// </summary>
        private static Queue<Bitmap> _imageList = new Queue<Bitmap>();
        /// <summary>
        /// 索引表
        /// </summary>
        private static Dictionary<Int32, IndexObject> _indexs = new Dictionary<Int32, IndexObject>();
        /// <summary>
        /// 每页等待时间
        /// </summary>
        private static Int32 deplyTime = 3000;

        private static PowerPoint.Application _objApp;
        private static PowerPoint._Presentation _objPres;

        private static Int32 width = Int32.Parse(ConfigurationManager.AppSettings["width"]);
        private static Int32 height = Int32.Parse(ConfigurationManager.AppSettings["height"]);

        /// <summary>
        /// 图片全局索引
        /// </summary>
        private static Int32 _index = 1;
        /// <summary>
        /// 图片前一个全局索引
        /// </summary>
        private static Int32 _prev_index = 1;

        // 图片的流水号
        private static Int32 _view_index = 1;
        // 图片对应的PPT页面编号
        private static Int32 _slide_index = -1;
        // 图片对应的前一个PPT页面编号
        private static Int32 _prev_slide_index = 1;

        /// <summary>
        /// 每秒图片张数
        /// </summary>
        //private static Int32 ImageNum = 20;

        private static String _noteFilePath;
        private static String _resultFilePath;
        private static String _outputPath;
        private static String _filePath;

        /// <summary>
        /// 索引对象
        /// </summary>
        private class IndexObject
        {
            /// <summary>
            /// PPT页索引
            /// </summary>
            public Int32 SlideIndex { get; set; }
            /// <summary>
            /// PPT页动画索引
            /// </summary>
            public Int32 ViewIndex { get; set; }
        }

        private static Int32 PixelsToPoints(Int32 val, Boolean vert)
        {
            Double result;
            if (vert)
            {
                result = Math.Truncate(val * 0.75);
            }
            else
            {
                result = Math.Truncate(val * 0.75);
            }

            return (Int32)result;
        }
        
        static void Main(string[] args)
        {
            ParseArgs(args);

            var fileName = Path.GetFileNameWithoutExtension(_filePath);

            _noteFilePath =  Path.Combine(_outputPath, fileName + "_note.txt");
            _resultFilePath = Path.Combine(_outputPath, fileName + "_result.txt");

            if (!Directory.Exists(Path.Combine(_outputPath, "images")))
            {
                Directory.CreateDirectory(Path.Combine(_outputPath, "images"));
            }
            
            var imgsPath = Path.Combine(_outputPath, "images.jpg");
            try
            {
                _objApp = new PowerPoint.Application();
                _objPres = _objApp.Presentations.Open(_filePath, WithWindow: MsoTriState.msoFalse);

                _objPres.SaveAs(imgsPath, PowerPoint.PpSaveAsFileType.ppSaveAsPNG, MsoTriState.msoTrue);

                
                // 保存备注
                SaveNoteInfo();

                // 设置呈现窗口
                SetSildeShowWindow();
            
                var pptRenderThread = InitPPTRenderThread();
                var imageSaveThread = InitImageSaveThread();

                pptRenderThread.Start();
                imageSaveThread.Start();

                while (_objApp.SlideShowWindows.Count >= 1) Thread.Sleep(deplyTime / 10);
            
                pptRenderThread.Abort();
                //pptRenderThread.Join();
                imageSaveThread.Abort();
                //imageSaveThread.Join();
                
                _objPres.Save();

                _objApp.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            // 杀死powerpnt
            KillProcess("powerpnt");
            
            SaveResult();

            // 退出
            Environment.Exit(0);
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
        /// 初始化参数
        /// </summary>
        private static void ParseArgs(String[] args)
        {
            // 文件路径
            _filePath = args[1];
            // 输出路径
            _outputPath = args[3];
        }

        /// <summary>
        /// 保存结果
        /// </summary>
        private static void SaveResult()
        {
            var stream = new FileStream(_resultFilePath, FileMode.Create);
            var writer = new StreamWriter(stream);

            var duration = 3;
            var prevSlideIndex = 0;

            // key是图片节点信息
            foreach (var key in _indexs.Keys)
            {
                if (_indexs[key].SlideIndex != prevSlideIndex)
                {
                    duration = 3;

                    // 前一个key与后一个key之间的图片为一组
                    writer.WriteLine(key + ";" + _indexs[key].SlideIndex + "-" + _indexs[key].ViewIndex + "@" + duration);
                }
                else
                {
                    // 前一个key与后一个key之间的图片为一组
                    writer.WriteLine(key + ";" + _indexs[key].SlideIndex + "-" + _indexs[key].ViewIndex + "@" + duration);

                    duration += 3;
                }
                

                writer.Flush();
                stream.Flush();
            }

            writer.Close();
            stream.Close();

            writer.Dispose();
            stream.Dispose();
        }

        /// <summary>
        /// 初始化图片保存线程
        /// </summary>
        /// <returns></returns>
        private static Thread InitImageSaveThread()
        {
            var imageIndex = 1;
            var imageSaveThread = new Thread(() =>
                {
                    while (true)
                    {
                        if (_imageList.Count == 0) continue;
                        lock (obj)
                        {
                            var bmp = _imageList.Dequeue();
                            using (var stream = new FileStream(Path.Combine(_outputPath, imageIndex.ToString("0000") + ".jpg"), FileMode.Create))
                            {
                                bmp.Save(stream, ImageFormat.Jpeg);

                                bmp.Dispose();
                            }

                            imageIndex++; 
                        }
                    }
                });

            return imageSaveThread;
        }


        /// <summary>
        /// 初始化抓图线程
        /// </summary>
        /// <param name="objPres"></param>
        /// <returns></returns>
        private static Thread InitCaptureThread(PowerPoint._Presentation objPres)
        {
            var thread = new Thread(() =>
            {
                while (true)
                {
                    try
                    {
                        var bmp = new Bitmap(width, height);
                        Graphics gSrc = Graphics.FromHwnd((IntPtr)objPres.SlideShowWindow.HWND);
                        IntPtr hdcSrc = gSrc.GetHdc();
                        Graphics gDes = Graphics.FromImage(bmp);
                        IntPtr hdcDes = gDes.GetHdc();

                        BitBlt(hdcDes, 0, 0, width, height, hdcSrc, 0, 0,
                               (int)(CopyPixelOperation.SourceCopy));

                        gSrc.ReleaseHdc(hdcSrc);
                        gDes.ReleaseHdc(hdcDes);
                        gSrc.Dispose();
                        gDes.Dispose();

                        _imageList.Enqueue(bmp);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    
                    // 图片索引+1
                    _index++;

                    Thread.Sleep(5);
                }
            });

            // 设置最大优先级
            thread.Priority = ThreadPriority.Highest;

            return thread;
        }

        /// <summary>
        /// 初始化PPT呈现线程
        /// </summary>
        /// <returns></returns>
        private static Thread InitPPTRenderThread()
        {
            var pptRenderThread = new Thread(() =>
                {
                    // 第一次
                    Thread thread = InitCaptureThread(_objPres);

                    try
                    {
                        if (_slide_index == -1)
                        {
                            _slide_index = 1;
                            _objPres.SlideShowWindow.View.First();
                            _objPres.SlideShowWindow.View.GotoSlide(1);
                            _objPres.SlideShowWindow.View.ResetSlideTime();

                            // 前一个全局索引等于_index
                            _prev_index = _index;

                            _indexs.Add(_index, new IndexObject()
                                {
                                    SlideIndex = _slide_index,
                                    ViewIndex = _view_index
                                });

                            thread.Start();
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }

                    while (true)
                    {
                        try
                        {
                            // 等待3秒
                            Thread.Sleep(deplyTime);

                            if (thread.ThreadState == ThreadState.Running
                                || thread.ThreadState != ThreadState.Unstarted)
                            {
                                _view_index++;
                                
                                thread.Abort();
                                thread = InitCaptureThread(_objPres);

                                // 前一个全局索引等于_index
                                _prev_index = _index;

                                // 启动之前，保存索引信息
                                _indexs.Add(_index, new IndexObject()
                                {
                                    SlideIndex = _slide_index,
                                    ViewIndex = _view_index
                                });

                                thread.Start();
                            }

                            // 下一个动画页面
                            _objPres.SlideShowWindow.View.Next();

                            if (_slide_index != -1)
                            {
                                // 获取当前PPT索引,第几张PPT
                                _slide_index = _objPres.SlideShowWindow.View.Slide.SlideIndex;

                                if (_prev_slide_index != _slide_index)
                                {
                                    _view_index = 1;
                                    _prev_slide_index = _slide_index;

                                    // 修改前一个全局索引对应的值
                                    _indexs[_prev_index] = new IndexObject()
                                        {
                                            SlideIndex = _slide_index,
                                            ViewIndex = _view_index
                                        };
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                    }
                });

            pptRenderThread.Priority = ThreadPriority.Highest;
            
            return pptRenderThread;
        }

        /// <summary>
        /// 设置呈现窗口
        /// </summary>
        private static void SetSildeShowWindow()
        {
            // 起始页面索引
            _objPres.SlideShowSettings.StartingSlide = 1;
            // 结束页面索引
            _objPres.SlideShowSettings.EndingSlide = _objPres.Slides.Count;
            // 放映开始
            _objPres.SlideShowSettings.Run();

            // 设置放映窗口宽度
            _objPres.SlideShowWindow.Width = PixelsToPoints(width, false);
            // 设置放映窗口高度
            _objPres.SlideShowWindow.Height = PixelsToPoints(height, true);
            // 设置放映窗口Top
            _objPres.SlideShowWindow.Top = 0;
            // 设置放映窗口Left
            _objPres.SlideShowWindow.Left = 0;
        }

        /// <summary>
        /// 保存备注信息
        /// </summary>
        private static void SaveNoteInfo()
        {
            var stream = new FileStream(_noteFilePath, FileMode.Create);
            var writer = new StreamWriter(stream);

            var note = new StringBuilder("{");

            foreach (PowerPoint.Slide slide in _objPres.Slides)
            {
                //writer.Write(slide.SlideIndex + "@");
                //读取每个幻灯片的备注文本  
                note.Append("{\"page\":\"" + slide.SlideIndex + "\",");

                foreach (PowerPoint.Shape nodeshape in slide.NotesPage.Shapes)
                {
                    if (nodeshape.TextFrame.HasText.Equals(MsoTriState.msoTrue))
                    {
                        //备注      
                        note.Append("\"note\":\"" + nodeshape.TextFrame.TextRange.Text + "\"");          
                        //writer.Write(nodeshape.TextFrame.TextRange.Text);
                    }
                }

                note.Append("},");
                //writer.WriteLine();
            }

            note = note.Remove(note.Length - 1, 1).Append("}");

            writer.Write(note);
            
            writer.Close();
            stream.Close();

            writer.Dispose();
            stream.Dispose();
        }
    }
}
