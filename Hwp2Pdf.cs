using System.Runtime.InteropServices;
using System.Text;
using HwpObjectLib;
using Microsoft.Win32;

#pragma warning disable CA1416
namespace DocsConverter
{
  public class Hwp2Pdf
  {
    [DllImport("kernel32", CharSet = CharSet.Unicode)]
    private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
    [DllImport("kernel32", CharSet = CharSet.Unicode)]
    private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

    private static readonly object lockObject = new();
    private static int st_convert_target_index;
    private static string[] target_type_array = new string[] { "PDF", "HWP", "HWPX", "HWPML2X", "HTML+", "ODT", "OOXML", "UNICODE", "RTF", "PNG" };
    private static string[] target_ext_array = new string[] { ".pdf", ".hwp", ".hwpx", ".hml", ".html", ".odt", ".docx", ".txt", ".rtf", ".png" };
    public static string[] source_ext_array = new string[] { ".hwp", ".hwpx", ".hml", ".html", ".odt", ".docx", ".doc", ".txt", ".rtf" };

    private string m_strPrinter = ""; // PDF 변환용 가상 프린터 이름
    private int m_nPrintMethod = 1;
    private string m_strSavePath = "";
    private int option_source_ext_flag = (1 | 2 | 4); // Source 확장자에 대한 비트플래그 타입
    private bool option_PDF_print; // true 면 가상인쇄 방식 사용
    private HwpObject hwp_object; // 한컴 오토메이션을 위한 기본 인터페이스
    private bool filecheckdll_ok;

    public Hwp2Pdf()
    {
      // 한컴오피스 설치여부 확인
      RegistryKey? reg = Registry.CurrentUser.OpenSubKey("SOFTWARE", true)?.OpenSubKey("HNC", true);
      if (reg == null)
      {
        Console.WriteLine("한컴오피스 한글2010 이상 버전이 설치되어 있지 않습니다.", "hwp2pdf");
        throw new Exception("한컴오피스 한글2010 이상 버전이 설치되어 있지 않습니다.");
      }

      try
      {
        hwp_object = new HwpObject(); // 한컴 오토메이션을 위한 기본 인터페이스
        if (hwp_object == null)
        {
          Console.WriteLine("한컴오피스 초기화에 실패했습니다. \r\n(알수 없는 이유)", "hwp2pdf");
          throw new Exception("한컴오피스 초기화에 실패했습니다. \r\n(알수 없는 이유)");
        }
      }
      catch (Exception ex)
      {
        Console.WriteLine($"한컴오피스 초기화에 실패했습니다. \r\n {ex.Message}", "hwp2pdf");
        throw new Exception($"한컴오피스 초기화에 실패했습니다. \r\n {ex.Message}");
      }

      // ini 파일에서 정보 불러오기
      string ini_path = AppDomain.CurrentDomain.BaseDirectory + "..\\..\\.." + "\\hwp2pdf.ini";
      StringBuilder strTemp = new(1024, 1024);
      GetPrivateProfileString("Main", "SavePath", "", strTemp, strTemp.Capacity, ini_path);
      m_strSavePath = strTemp.ToString();
      GetPrivateProfileString("Main", "OptionExtFlags", "", strTemp, strTemp.Capacity, ini_path);  // 과거 버전은 OptionExtFlag 
      if (strTemp.Length > 0) option_source_ext_flag = int.Parse(strTemp.ToString());
      GetPrivateProfileString("Main", "OptionPDFPrint", "", strTemp, strTemp.Capacity, ini_path);
      if (strTemp.Length > 0) option_PDF_print = bool.Parse(strTemp.ToString());
      GetPrivateProfileString("Main", "PrinterName", "", strTemp, strTemp.Capacity, ini_path);
      m_strPrinter = strTemp.ToString();
      GetPrivateProfileString("Main", "PrintMethod", "", strTemp, strTemp.Capacity, ini_path);
      if (strTemp.Length > 0) m_nPrintMethod = int.Parse(strTemp.ToString());

      // 레지스트리에 보안모듈 추가
      reg = reg.CreateSubKey("HwpAutomation");
      if (reg == null)
      {
        Console.WriteLine("레지스트리 항목(HwpAutomation) 추가 중 오류가 발생했습니다.", "hwp2pdf");
        throw new Exception("레지스트리 항목(HwpAutomation) 추가 중 오류가 발생했습니다.");
      }
      reg = reg.CreateSubKey("Modules");
      if (reg == null)
      {
        Console.WriteLine("레지스트리 항목(Modules) 추가 중 오류가 발생했습니다.", "hwp2pdf");
        throw new Exception("레지스트리 항목(Modules) 추가 중 오류가 발생했습니다.");
      }
      string dll_path = AppDomain.CurrentDomain.BaseDirectory + "..\\..\\..";
      dll_path += "\\" + "FilePathCheckerModuleExample.dll";
      FileInfo file_info = new(dll_path);
      if (file_info.Exists)
      {
        reg.SetValue("FilePathCheckerModuleExample", dll_path);
      }
      else
      {
        Console.WriteLine("실행파일과 같은 경로에 FilePathCheckerModuleExample.DLL이 없습니다.");
      }

      filecheckdll_ok = hwp_object.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample");
      if (filecheckdll_ok == false)
      {
        Console.WriteLine("FilePathCheckerModuleExample.DLL 연결에 실패했습니다.");
      }

      // PDF 변환용 프린터가 설치되어 있는지 확인
      System.Collections.ArrayList printer_names = new(System.Drawing.Printing.PrinterSettings.InstalledPrinters);
      bool bPrinterInstalled = false;
      bool bSetDefault = false;
      if (m_strPrinter == "") bSetDefault = true;
      for (int i = 0; i < printer_names.Count; i++)
      {
        string? name = printer_names[i]?.ToString();
        if (name != null && name == m_strPrinter)
        {
          bPrinterInstalled = true;
          break;
        }
        else if (name != null && name.ToLower().Contains("pdf"))
        {
          if (name.ToLower().Contains("hancom"))
          {
            bPrinterInstalled = true;
            if (bSetDefault == true)
            {   // 프린터가 설정되지 않은 상태에서 한컴PDF가 있으면 기본 프린터로 사용 
              m_strPrinter = name;
              break;
            }
          }
          if (name.ToLower().Contains("microsoft"))
          {
            bPrinterInstalled = true;
            // 프린터가 설정되지 않은 상태에서 한컴PDF가 없는 경우 MS PDF 사용
            if (bSetDefault == true) m_strPrinter = name;
          }
        }
      }
      if (bPrinterInstalled == false)
      {
        Console.WriteLine("한컴 PDF 또는 Micosoft Print to PDF가 설치되어 있지 않습니다.", "hwp2pdf");
      }

      int nCount = Math.Min(target_ext_array.GetLength(0), target_type_array.GetLength(0));
    }

    public void Convert_file(string input_filename, string output_filename)
    {
      string basePath = AppDomain.CurrentDomain.BaseDirectory + @"\";
      string hwpPath = basePath + "files";
      string input_path = hwpPath + "\\" + input_filename;
      string output_path = hwpPath + "\\" + output_filename;

      string input_ext = System.IO.Path.GetExtension(input_filename).ToLower();
      string output_ext = System.IO.Path.GetExtension(output_filename).ToLower();

      int input_index = Array.IndexOf(source_ext_array, input_ext);
      if (input_index == -1)
      {
        throw new ArgumentException("지원하지 않는 입력 파일 확장자입니다.");
      }

      int output_index = Array.IndexOf(target_ext_array, output_ext);
      if (output_index == -1)
      {
        throw new ArgumentException("지원하지 않는 출력 파일 확장자입니다.");
      }

      st_convert_target_index = output_index;

      Thread th = new(() => Convert_thread(input_path, output_path));
      th.SetApartmentState(ApartmentState.STA);
      lock (lockObject)
      {
        th.Start();
        th.Join();
      }
      Console.WriteLine("변환 완료");
    }

    private void Convert_thread(string input_path, string output_path)
    {
      if (filecheckdll_ok == false)
      {
        IXHwpWindows hwp_windows = (IXHwpWindows)hwp_object.XHwpWindows;
        IXHwpWindow hwp_window = (IXHwpWindow)hwp_windows.Item[0];
        hwp_window.Visible = filecheckdll_ok;
      }
      Console.WriteLine("파일 변환을 시작합니다. 잠시 기다려 주세요......");
      string target_type = target_type_array[st_convert_target_index];
      string target_ext = target_ext_array[st_convert_target_index];
      if (filecheckdll_ok == true) hwp_object.SetMessageBoxMode(0x00211411); // HwpCtrl API 문서에 있음
      string file_ext = System.IO.Path.GetExtension(input_path).ToLower();
      if (file_ext == target_ext)
      {
        Console.WriteLine("변환안함(같은형식)");
        System.IO.File.Copy(input_path, output_path, true);
      }
      else if (hwp_object.Open(input_path, "", "lock:false;forceopen:true;suspendpassword:true;")) // 포맷을 지정하지 않아도 자동 인식
      {
        Console.WriteLine("변환중");
        bool bSuccess;
        if (target_type.ToUpper() == "PDF" && option_PDF_print == true && m_strPrinter != "")
        {
          HAction hwp_action = (HAction)hwp_object.HAction;
          HParameterSet hwp_pset = (HParameterSet)hwp_object.HParameterSet;
          HPrint hwp_print = (HPrint)hwp_pset.HPrint;
          HSet hwp_set = (HSet)hwp_print.HSet;
          hwp_action.GetDefault("Print", hwp_set);
          hwp_print.PrintMethod = (ushort)m_nPrintMethod;
          hwp_print.Collate = 1;
          hwp_print.NumCopy = 1;
          //hwp_print.UserOrder = 0;
          hwp_print.PrintToFile = 1;
          hwp_print.Range = 1;
          hwp_print.filename = output_path;
          // hwp_print.filename = "outout.png";
          // hwp_print.PrinterName = m_strPrinter;
          //hwp_print.UsingPagenum = 1;
          //hwp_print.ReverseOrder = 0;
          //hwp_print.Pause = 0;
          //hwp_print.PrintImage = 1;
          //hwp_print.PrintDrawObj = 1;
          //hwp_print.PrintClickHere = 0;
          //hwp_print.PrintFormObj = 1;
          //hwp_print.PrintMarkPen = 0;  // 추후 옵션 여부 검토할 것
          //hwp_print.PrintMemo = 0;
          //hwp_print.PrintMemoContents = 0;
          //hwp_print.PrintRevision = 1;
          //hwp_print.PrintBarcode = 1;
          // hwp_print.filename = "outout.png";
          // hwp_print.PrinterName = m_strPrinter;
          //hwp_print.UsingPagenum = 1;

          //hwp_print.PrintBarcode = 1;
          hwp_print.Flags = 8192;
          hwp_print.Device = 3;
          //hwp_print.PrintPronounce = 0;
          bSuccess = hwp_action.Execute("Print", hwp_set); //PrintToPDF를 쓰면 폰트에 따라 숫자, 첨자 등이 안나올수 있음
          if (bSuccess == true)
          {
            Console.WriteLine("파일 쓰는 중");
            //한컴 PDF Printer는 인쇄가 성공했더라도 실제 파일이 저장되었는지 확인도 필요함. 
            //별도 쓰레드가 돌아가면서 성공값이 리턴된 후에도 파일IO가 계속되는 경우가 있음
            //이렇게 하지 않으면 중간에 확인창이 뜬다.
            FileStream? stream = null;
            bool bWriteFinished = false;
            int elapsedTime = 0;
            while (bWriteFinished == false && elapsedTime < 30000) // 30초 제한
            {
              try
              {
                stream = new FileStream(output_path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                if (stream != null) bWriteFinished = true;
              }
              catch (IOException)
              {
                bWriteFinished = false; //쓰기가 끝나지 않은 상태면
                Thread.Sleep(500); // 0.5초 대기
                elapsedTime += 500;
              }
            }
            stream?.Close();
            if (bWriteFinished)
            {
              Console.WriteLine("쓰기 종료");
            }
            else
            {
              Console.WriteLine("파일 쓰기 실패");
              bSuccess = false;
            }
          }
        }

        if (target_type.ToUpper() == "PNG")
        {
          HAction hwp_action = (HAction)hwp_object.HAction;
          HParameterSet hwp_pset = (HParameterSet)hwp_object.HParameterSet;
          HPrint hwp_print = (HPrint)hwp_pset.HPrint;
          HSet hwp_set = (HSet)hwp_print.HSet;
          hwp_action.GetDefault("Print", hwp_set);
          hwp_print.PrintMethod = (ushort)m_nPrintMethod;
          hwp_print.Collate = 1;
          hwp_print.NumCopy = 1;
          hwp_print.PrintToFile = 1;
          hwp_print.Range = 1;
          hwp_print.filename = output_path;
          hwp_print.PrinterName = m_strPrinter;
          hwp_print.Flags = 8192;
          hwp_print.Device = 2;
          bSuccess = hwp_action.Execute("Print", hwp_set);
          if (bSuccess == true)
          {
            Console.WriteLine("파일 쓰는 중");
            FileStream? stream = null;
            bool bWriteFinished = false;
            int elapsedTime = 0;
            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string filePath = Path.Combine(documentsPath, "image001.png");
            while (bWriteFinished == false && elapsedTime < 10000) // 10초 제한
            {
              try
              {
                stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                if (stream != null) bWriteFinished = true;
              }
              catch (IOException)
              {
                bWriteFinished = false;
                Thread.Sleep(10); // 0.01초 대기
                elapsedTime += 1000;
              }
            }
            stream?.Close();
            if (bWriteFinished)
            {
              Console.WriteLine("쓰기 종료");
            }
            else
            {
              Console.WriteLine("파일 쓰기 실패");
              bSuccess = false;
            }
          }
        }
        else
        {
          //SaveAs의 경우 PDF 변환시 HWP파일의 모아찍기 설정은 그대로 유지됨

          bSuccess = hwp_object.SaveAs(output_path, target_type, "");
        }
        if (bSuccess)
        {
          Console.WriteLine("완료(덮어씀)");
        }
        else Console.WriteLine("변환 시도 실패");
        hwp_object.Clear(1);
      }
      else Console.WriteLine("원본파일 열기 실패");
      Console.WriteLine(string.Format("파일을 변환하였습니다."));
    }
  }
}