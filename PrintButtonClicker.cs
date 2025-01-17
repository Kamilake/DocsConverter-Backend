using System.Runtime.InteropServices;

class PrintButtonClicker
{
  [DllImport("user32.dll", SetLastError = true)]
  static extern IntPtr FindWindow(string? lpClassName, string lpWindowName);

  [DllImport("user32.dll", SetLastError = true)]
  static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string? lpszClass, string lpszWindow);

  [DllImport("user32.dll", SetLastError = true)]
  static extern IntPtr GetDlgItem(IntPtr hDlg, int nIDDlgItem);

  [DllImport("user32.dll", CharSet = CharSet.Auto)]
  static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

  const uint BM_CLICK = 0x00F5;

  public Thread clickerThread = new(new ThreadStart(ClickButton));

  static void ClickButton()
  {
    while (true)
    {
      // 1. 윈도우 찾기
      IntPtr hWnd = FindWindow(null, "그림으로 저장하기");
      if (hWnd != IntPtr.Zero)
      {
        Console.WriteLine("윈도우를 찾았어요!");
        // 2. 버튼 찾기
        IntPtr hButton = FindWindowEx(hWnd, IntPtr.Zero, null, "저장(&S)");
        // hButton의 아이디 출력
        Console.WriteLine(hButton);
        if (hButton != IntPtr.Zero)
        {
          // 3. 버튼 클릭하기
          SendMessage(hButton, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
          Console.WriteLine("버튼을 클릭했어요!");
        }
      }
      Thread.Sleep(50);
    }
  }
}