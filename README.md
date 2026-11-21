# SeleniumBasic 3.2

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

3.1 and 3.2 add what current sites often need: Shadow DOM, downloads, PDF, waiting for the network, infinite scroll, tabs, cookies saved to a file, DevTools commands and more. See [New in 3.2](#new-in-32) and [New in 3.1](#new-in-31).

## Install

1. Close Excel, Word and any script using Selenium.
2. Run `SeleniumBasic-3.2.0.0.exe` from the [Releases](https://github.com/naesesang/SeleniumBasic/releases/latest) page. No administrator rights are needed: it installs for the current user in `%LOCALAPPDATA%\SeleniumBasic` and replaces a previous 2.0.9 or 3.x install.
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

## New in 3.2

```vba
Dim driver As New Selenium.ChromeDriver
driver.SetHeadless                              ' options: call before Start
driver.HideAutomationInfo
driver.Start

driver.Get "https://example.com/list"
driver.WaitForNetworkIdle                       ' until the fetch/XHR requests stop
n = driver.ScrollUntilEnd                       ' infinite scroll: load everything
Debug.Print driver.StatusCode                   ' 200, 404...

Set item = driver.FindElementByCss(".item")
item.Parent.Highlight                           ' outline an element while debugging
item.FindElementByTag("img").DownloadResource "C:\Temp\"   ' with the login cookies

driver.FindElementById("open").Click
driver.SwitchToNewWindow                        ' the tab opened by the click
driver.Manage.Cookies.SaveToFile "C:\Temp\login.json"     ' restore later with LoadFromFile

driver.Get "https://api.example.com/items"
Set data = driver.PageToJson                    ' Dictionary / List
Debug.Print data("items")(1)("name")
```

| Area | Added |
|---|---|
| Options (before `Start`) | `SetHeadless`, `SetIncognito`, `SetGuest`, `SetUserAgent`, `HideAutomationInfo`, `SetAcceptInsecureCerts`, `SetPageLoadStrategy`, `SetUnhandledPromptBehavior`, `SetDetach` (browser stays open after `Quit`), `EnableBiDi` + `WebSocketUrl` |
| Waits | `WaitForNetworkIdle`, `WaitForDomIdle` (driver and element), `WaitForJQueryIdle` |
| Scrolling (driver and element) | `ScrollToTop`, `ScrollToBottom`, `ScrollTo`, `ScrollBy`, `ScrollUntilEnd`, `ScrollHeight`, `ScrollWidth` |
| Page | `StatusCode`, `UserAgent`, `OpenHtml`, `OpenFile`, `PageToHtmlDocument` (MSHTML), `PageToJson`, `PageToXmlDocument` (MSXML), `DownloadResource` |
| Element | `Parent`, `Children`, `SetAttribute`, `RemoveAttribute`, `HasAttribute`, `Property`, `DomAttribute`, `Focus`, `HasFocus`, `AriaLabel`, `AriaRole`, `Highlight`, `Unhighlight` (also on `WebElements`), `UploadFile`, `DownloadResource` |
| Windows | `NewTab`, `NewWindow`, `SwitchToNewWindow`, `SwitchToWindowByUrl`, `Window.Minimize`, `Window.Url`, `Window.Handle` |
| Actions and keys | `Actions.ScrollByWheel`, `ScrollToElement`, `MoveToPosition`, `DragAndDropToPosition`, `Keys.Repeat`, `Keys.Chord` |
| Cookies | `Cookies.SaveToFile`, `Cookies.LoadFromFile`, `Cookie.SameSite`, `Cookie.HttpOnly`, `AddCookie(..., sameSite)` |
| DevTools and session | `ExecuteCDP` (Chrome, Edge), `SessionId`, `Status`, `CurrentFrameName`, `Manage.SetLocation` / `Location` (now works with current drivers) |
| PDF | `PrintToPDF` options: `paperSize` (A4, Letter, "21x29.7"...), `margin`, `shrinkToFit`, `displayHeaderFooter`, `headerTemplate`, `footerTemplate` |
| Utils (`New Selenium.Utils`) | `ParseJson`, `ToJson`, `SaveBase64ToFile`, `FileToBase64` |

Each command shows its description in the VBA Object Browser (F2). A command reference in Korean is in [docs/COMMANDS.ko.md](docs/COMMANDS.ko.md).

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

### 3.2 새 기능

```vba
driver.SetHeadless                              ' 창 없이 실행 (Start 전에 호출)
driver.HideAutomationInfo                       ' "자동화된 소프트웨어가 제어 중" 알림 숨김
driver.Start

driver.WaitForNetworkIdle                       ' 페이지의 네트워크 요청이 끝날 때까지 대기
n = driver.ScrollUntilEnd                       ' 무한 스크롤을 끝까지 내리기
Debug.Print driver.StatusCode                   ' 200, 404 ...

Set item = driver.FindElementByCss(".item")
item.Parent.Highlight                           ' 부모 요소를 빨간 테두리로 표시
item.FindElementByTag("img").DownloadResource "C:\Temp\"   ' 로그인 상태 그대로 이미지 저장

driver.FindElementById("open").Click
driver.SwitchToNewWindow                        ' 클릭으로 새로 열린 탭으로 전환
driver.Manage.Cookies.SaveToFile "C:\Temp\login.json"     ' 다음에 LoadFromFile로 로그인 복원

driver.Get "https://api.example.com/items"
Set data = driver.PageToJson                    ' JSON 응답을 Dictionary/List로
```

**모든 명령의 한국어 설명은 [명령어 설명서](docs/COMMANDS.ko.md)에 있습니다.** 추가된 명령 목록은 위의 [New in 3.2](#new-in-32) 표에도 있습니다. VBA 편집기에서 `driver.`를 입력하면 자동 완성 목록에 나오고, 개체 찾아보기(F2)에서 설명을 볼 수 있습니다.

- 브라우저 설정(`SetHeadless`, `SetIncognito`, `SetUserAgent` 등)은 `Start` 전에 호출해야 합니다.
- `SetDetach`를 쓰면 `Quit` 후에도 브라우저 창이 남아 있습니다.
- `ExecuteCDP`, 머리글·바닥글이 있는 PDF, 위치 설정은 Chrome과 Edge에서만 동작합니다.

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
2. [Releases](https://github.com/naesesang/SeleniumBasic/releases/latest)에서 `SeleniumBasic-3.2.0.0.exe`를 받아 실행합니다. 관리자 권한은 필요 없고, 기존 2.0.9나 3.x는 자동으로 교체됩니다.
3. 코드 서명이 없어 SmartScreen 경고가 뜰 수 있습니다. "추가 정보" > "실행"을 누르세요. 스마트 앱 컨트롤이 켜진 PC에서는 이 선택지 없이 차단됩니다.

필요 환경: Windows 10/11 (.NET Framework 4.8 기본 포함), Chrome 또는 Edge, 처음 실행할 때 인터넷 연결.

### 문제 해결

- **회사망 등에서 드라이버 다운로드가 막히는 경우:** 브라우저 버전에 맞는 `chromedriver.exe`(또는 `msedgedriver.exe`)를 직접 받아 `%LOCALAPPDATA%\SeleniumBasic` 폴더의 `Selenium.dll` 옆에 넣으세요.
- **런타임 오류 8013101B / 80131040:** 엑셀을 닫고 설치 파일을 다시 실행하세요.
- 오류 기록이 필요하면 `%LOCALAPPDATA%\SeleniumBasic\debug.log` 빈 파일을 만든 뒤 오류를 재현하세요.

문의와 버그 제보는 [Issues](https://github.com/naesesang/SeleniumBasic/issues)에 남겨 주세요. 별도의 지원을 보장하지 않으며 있는 그대로 제공됩니다.
