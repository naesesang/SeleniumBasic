# SeleniumBasic 3.1

Browser automation for VBA (Excel, Access, Word), VBScript and VB.Net that works with **current Chrome and Edge**.

This is an unofficial, updated fork of [florentbr/SeleniumBasic](https://github.com/florentbr/SeleniumBasic) 2.0.9, which has not been updated since 2016. It is not affiliated with or endorsed by the original author.

**[Download the installer (Releases)](https://github.com/naesesang/SeleniumBasic/releases/latest)**

> 한국어 안내는 [아래](#한국어)에 있습니다.

## Why this fork

SeleniumBasic 2.0.9 stopped working with current browsers:

- The driver must match the browser's major version, so every Chrome update breaks it with errors such as `SessionNotCreatedException: This version of ChromeDriver only supports Chrome version ...`
- Current chromedriver and msedgedriver only speak the W3C WebDriver protocol, which 2.0.9 does not support.

3.0 fixes this:

- **Automatic driver download:** the driver matching the installed Chrome or Edge is downloaded on first use. Browser updates no longer require a manual driver swap.
- **W3C WebDriver protocol.**
- **Your existing VBA code keeps working.** The type library is unchanged, so the "Selenium Type Library" reference and code such as `New Selenium.ChromeDriver` stay the same.

3.1 adds what current sites often need: **Shadow DOM**, **download folder and waiting for downloads**, and **saving a page as PDF**. See [New in 3.1](#new-in-31).

## Install

1. Close Excel, Word and any script using Selenium.
2. Run `SeleniumBasic-3.1.0.0.exe` from the [Releases](https://github.com/naesesang/SeleniumBasic/releases/latest) page. No administrator rights are needed: it installs for the current user in `%LOCALAPPDATA%\SeleniumBasic` and replaces a previous 2.0.9 or 3.0 install.
3. In the VBA editor, check *Tools > References > Selenium Type Library* (already checked in workbooks made with 2.0.9).

The installer is not code-signed, so SmartScreen may show a warning ("More info" > "Run anyway"). On PCs where Smart App Control is on, Windows blocks unsigned installers without that option.

Requirements: Windows 10/11 with the .NET Framework 4.8 (built into Windows 10 1903+ and Windows 11), Chrome or Edge, and internet access on first use of each browser version.

## Example

```vba
Sub Example()
    Dim driver As New Selenium.ChromeDriver   ' or Selenium.EdgeDriver
    driver.Start
    driver.Get "https://www.example.com"
    Debug.Print driver.Title                  ' Example Domain
    Debug.Print driver.FindElementByCss("p").Text
    driver.Quit
End Sub
```

More examples are in [Examples](Examples).

## New in 3.1

```vba
' Shadow DOM: find the elements inside a web component
Set root = driver.FindElementByCss("my-component").ShadowRoot
root.FindElementByCss("button.save").Click

' Downloads: choose the folder, then wait for the file
driver.SetDownloadFolder "C:\Temp\Reports"      ' before or after Start (Firefox: before)
driver.FindElementById("export").Click
path = driver.WaitForDownload("*.xlsx")         ' full path of the file, 60 s timeout by default
Workbooks.Open path

' Save the page as PDF, as printed by the browser
driver.PrintToPDF "C:\Temp\page.pdf"
driver.PrintToPDF "C:\Temp\wide.pdf", landscape:=True, scale:=0.8, pageRanges:="1-2"
```

- `ShadowRoot` waits for the shadow root like `FindElement` (optional `timeout` and `raise`). Inside a shadow root, use CSS, id, name, class or tag; Chrome and Edge do not support XPath or link text there.
- `SetDownloadFolder` saves downloads without asking and downloads PDF files instead of opening them.
- `WaitForDownload` waits for the next file downloaded after `Start`, `SetDownloadFolder` or the previous `WaitForDownload`, and returns its path. Without `SetDownloadFolder` it watches the user's Downloads folder.

## Drivers

Drivers are cached in `%LOCALAPPDATA%\SeleniumBasic\drivers`. If the download is blocked (offline PC, proxy requiring authentication), download the driver matching your browser version yourself and put it next to `Selenium.dll` in `%LOCALAPPDATA%\SeleniumBasic`:

- Chrome: `chromedriver.exe` from [Chrome for Testing](https://googlechromelabs.github.io/chrome-for-testing/)
- Edge: `msedgedriver.exe` from [Microsoft Edge WebDriver](https://developer.microsoft.com/microsoft-edge/tools/webdriver/)

## Changes since 2.0.9

See [CHANGELOG.txt](CHANGELOG.txt). In short (3.0): .NET Framework 4.8 (was 3.5), W3C protocol, driver auto-download, Edge on msedgedriver, Firefox on geckodriver, TLS 1.2. PhantomJS, Opera, the IE driver and the Selenium IDE add-on are removed.

## Troubleshooting

- **Runtime error 8013101B or 80131040** when creating a driver: the COM registration still points to an older version. Close Excel and run the installer again.
- **`Failed to get chromedriver.exe ...`**: the driver download failed. See [Drivers](#drivers).
- To collect a diagnostic log, create an empty file `%LOCALAPPDATA%\SeleniumBasic\debug.log` and reproduce the error.

Tested with Chrome and Edge 156 to 158, VBA in 64-bit Excel (Microsoft 365) and VBScript 32/64-bit. Not tested yet: Firefox, 32-bit Office. Reports are welcome in [Issues](https://github.com/naesesang/SeleniumBasic/issues).

This project is provided as is, without warranty or guaranteed support.

## Build

Requires the .NET SDK and [Inno Setup 6](https://jrsoftware.org/isinfo.php):

```powershell
.\build-setup.ps1
```

## License

BSD 3-Clause, see [LICENSE.txt](LICENSE.txt). Original work copyright (c) 2015 florentbr.

---

## 한국어

최신 Chrome과 Edge에서 동작하도록 업데이트한 SeleniumBasic입니다. 원작([florentbr/SeleniumBasic](https://github.com/florentbr/SeleniumBasic) 2.0.9)과는 무관한 비공식 포크입니다.

**크롬 업데이트 때마다 드라이버를 교체할 필요가 없습니다.** 설치된 Chrome/Edge 버전에 맞는 드라이버를 처음 실행할 때 자동으로 내려받습니다. `This version of ChromeDriver only supports Chrome version ...` 같은 오류로 막혔다면 이 버전을 설치하세요.

**기존 VBA 코드는 그대로 쓸 수 있습니다.** 참조(Selenium Type Library)와 `New Selenium.ChromeDriver` 같은 코드를 바꿀 필요가 없습니다.

### 3.1 새 기능

```vba
' Shadow DOM 안쪽 요소 찾기
Set root = driver.FindElementByCss("my-component").ShadowRoot
root.FindElementByCss("button.save").Click

' 다운로드 폴더 지정 후 다운로드 완료까지 대기
driver.SetDownloadFolder "C:\Temp\Reports"      ' Start 전후 모두 가능 (Firefox는 Start 전)
driver.FindElementById("export").Click
path = driver.WaitForDownload("*.xlsx")         ' 받은 파일의 전체 경로, 기본 60초 대기
Workbooks.Open path

' 페이지를 PDF로 저장
driver.PrintToPDF "C:\Temp\page.pdf"
```

- Shadow DOM 안에서는 CSS, id, name, class, 태그로 찾으세요. Chrome과 Edge는 그 안에서 XPath와 링크 텍스트 검색을 지원하지 않습니다.
- `SetDownloadFolder`를 쓰면 저장 위치를 묻지 않고 저장하며, PDF는 브라우저에서 열지 않고 내려받습니다.
- `WaitForDownload`는 `Start`, `SetDownloadFolder`, 이전 `WaitForDownload` 이후에 받은 파일을 기다립니다. 폴더를 지정하지 않으면 사용자의 다운로드 폴더를 확인합니다.

### 설치

1. 엑셀과 Selenium을 쓰는 프로그램을 모두 닫습니다.
2. [Releases](https://github.com/naesesang/SeleniumBasic/releases/latest)에서 `SeleniumBasic-3.1.0.0.exe`를 받아 실행합니다. 관리자 권한은 필요 없고, 기존 2.0.9나 3.0은 자동으로 교체됩니다.
3. 코드 서명이 없어 SmartScreen 경고가 뜰 수 있습니다. "추가 정보" > "실행"을 누르세요. 스마트 앱 컨트롤이 켜진 PC에서는 이 선택지 없이 차단됩니다.

필요 환경: Windows 10/11 (.NET Framework 4.8 기본 포함), Chrome 또는 Edge, 처음 실행할 때 인터넷 연결.

### 문제 해결

- **회사망 등에서 드라이버 다운로드가 막히는 경우:** 브라우저 버전에 맞는 `chromedriver.exe`(또는 `msedgedriver.exe`)를 직접 받아 `%LOCALAPPDATA%\SeleniumBasic` 폴더의 `Selenium.dll` 옆에 넣으세요.
- **런타임 오류 8013101B / 80131040:** 엑셀을 닫고 설치 파일을 다시 실행하세요.
- 오류 기록이 필요하면 `%LOCALAPPDATA%\SeleniumBasic\debug.log` 빈 파일을 만든 뒤 오류를 재현하세요.

문의와 버그 제보는 [Issues](https://github.com/naesesang/SeleniumBasic/issues)에 남겨 주세요. 별도의 지원을 보장하지 않으며 있는 그대로 제공됩니다.
