# SeleniumBasic 3.2 명령어 설명서 (3.2.1 기준)

VBA(Excel, Access, Word)와 VBScript에서 쓸 수 있는 SeleniumBasic의 명령을 객체별로 정리한 설명서입니다.
3.1과 3.2에서 추가된 명령에는 **3.1**, **3.2** 표시를 붙였습니다.

## 목차

1. [시작하기](#1-시작하기)
2. [공통 규칙](#2-공통-규칙)
3. [WebDriver: 브라우저](#3-webdriver-브라우저)
4. [WebElement: 요소](#4-webelement-요소)
5. [WebElements: 요소 목록](#5-webelements-요소-목록)
6. [ShadowRoot: Shadow DOM](#6-shadowroot-shadow-dom)
7. [By: 찾는 방법](#7-by-찾는-방법)
8. [Window: 창](#8-window-창)
9. [Manage, Cookies, Cookie, Storage: 쿠키와 저장소](#9-manage-cookies-cookie-storage-쿠키와-저장소)
10. [Timeouts: 대기 시간](#10-timeouts-대기-시간)
11. [Alert: 알림창](#11-alert-알림창)
12. [Actions, Mouse, Keyboard, Keys: 마우스와 키보드](#12-actions-mouse-keyboard-keys-마우스와-키보드)
13. [SelectElement, TableElement: 선택 상자와 표](#13-selectelement-tableelement-선택-상자와-표)
14. [Utils, Waiter: 도구](#14-utils-waiter-도구)
15. [Image, PdfFile: 이미지와 PDF](#15-image-pdffile-이미지와-pdf)
16. [List, Dictionary, Table: 데이터](#16-list-dictionary-table-데이터)
17. [Assert, Verify: 검증](#17-assert-verify-검증)
18. [Proxy, Point, Size, Application](#18-proxy-point-size-application)
19. [현재 드라이버에서 동작하지 않는 명령](#19-현재-드라이버에서-동작하지-않는-명령)

---

## 1. 시작하기

VBA 편집기에서 **도구 > 참조 > Selenium Type Library**에 체크한 뒤 사용합니다.

```vba
Sub 시작()
    Dim driver As New Selenium.ChromeDriver    ' Edge는 Selenium.EdgeDriver
    driver.Start                               ' 맞는 드라이버를 자동으로 내려받고 브라우저 실행
    driver.Get "https://www.example.com"
    Debug.Print driver.Title
    Debug.Print driver.FindElementByCss("p").Text
    driver.Quit                                ' 브라우저 종료
End Sub
```

**만들 수 있는 객체** (`New Selenium.XXX` 또는 `CreateObject("Selenium.XXX")`)

| 객체 | 용도 |
|---|---|
| `ChromeDriver`, `EdgeDriver`, `FirefoxDriver` | 브라우저별 드라이버 |
| `WebDriver` | `Start "chrome"`처럼 브라우저를 나중에 정하는 드라이버 |
| `By` | 요소 찾는 방법 ([7장](#7-by-찾는-방법)) |
| `Keys` | 특수 키 ([12장](#12-actions-mouse-keyboard-keys-마우스와-키보드)) |
| `Utils`, `Waiter` | 정규식, JSON, Base64, 대기 도구 |
| `Assert`, `Verify` | 검증 |
| `List`, `Dictionary`, `Table` | 데이터 묶음 |
| `Image`, `PdfFile` | 이미지 비교, PDF 만들기 |

---

## 2. 공통 규칙

- **시간 단위는 밀리초**입니다. `1000` = 1초.
- **`timeout` 인수**
  - 요소를 찾는 명령은 `-1`(기본값)이면 `Timeouts.ImplicitWait`(기본 3초)만큼 기다립니다.
  - `0`이면 기다리지 않습니다.
- **`raise` 인수**
  - `True`(기본값)면 못 찾았을 때 오류를 냅니다.
  - `False`면 오류 대신 `Nothing`(또는 `False`)을 돌려줍니다.

  ```vba
  Set el = driver.FindElementById("popup", 2000, False)
  If el Is Nothing Then Debug.Print "팝업 없음"
  ```
- **번호는 1부터** 셉니다. `WebElements(1)`, `List(1)`, `Cookies(1)`이 첫 번째 항목입니다.
- **경로**에는 `%USERPROFILE%` 같은 환경 변수를 쓸 수 있습니다.
- **선택 인수는 생략할 수 있습니다.** 이름을 붙여 일부만 지정할 수도 있습니다.

  ```vba
  driver.PrintToPDF "C:\a.pdf", paperSize:="A4"
  ```
- **브라우저 설정 명령**(`SetHeadless` 등, **3.2**)은 `Start` **전에** 호출해야 합니다.

---

## 3. WebDriver: 브라우저

`ChromeDriver`, `EdgeDriver`, `FirefoxDriver`, `WebDriver`는 모두 아래 명령을 가집니다.

### 3.1 실행과 종료

| 명령 | 설명 |
|---|---|
| `Start([browser], [baseUrl])` | 브라우저를 실행합니다. `browser`는 `WebDriver`에서만 지정합니다: `"chrome"`, `"edge"`, `"firefox"` |
| `StartRemotely(executorUri, [browser], [version], [platform])` | Selenium Grid 같은 원격 서버에서 브라우저를 실행합니다. 예: `"http://localhost:4444/wd/hub"` |
| `Quit()` | 브라우저와 드라이버를 종료합니다. |
| `Close()` | 현재 창(탭)만 닫습니다. |
| `SessionId` **3.2** | 세션 ID |
| `Status()` **3.2** | 드라이버 상태(`ready`, `message` 등)를 담은 Dictionary |

### 3.2 브라우저 설정 (`Start` 전에 호출)

| 명령 | 설명 |
|---|---|
| `SetHeadless([enable])` **3.2** | 창 없이 실행합니다. |
| `SetIncognito([enable])` **3.2** | 시크릿 모드(Edge는 InPrivate, Firefox는 사생활 보호 모드)로 실행합니다. |
| `SetGuest([enable])` **3.2** | 게스트 모드로 실행합니다(Chrome, Edge). |
| `SetUserAgent(userAgent)` **3.2** | 브라우저가 보내는 User-Agent를 바꿉니다. |
| `HideAutomationInfo()` **3.2** | "자동화된 테스트 소프트웨어에 의해 제어되고 있습니다" 알림을 숨깁니다(Chrome, Edge). |
| `SetAcceptInsecureCerts([accept])` **3.2** | 인증서 오류가 있는 사이트도 엽니다. |
| `SetPageLoadStrategy(strategy)` **3.2** | `Get`이 언제 돌아올지 정합니다. `"normal"`은 전체 로딩 후(기본값), `"eager"`는 문서 준비 후, `"none"`은 기다리지 않음입니다. |
| `SetUnhandledPromptBehavior(behavior)` **3.2** | 예상치 못한 알림창 처리 방식입니다: `"dismiss"`, `"accept"`, `"dismiss and notify"`, `"accept and notify"`, `"ignore"` |
| `SetDetach([detach])` **3.2** | `Quit` 후나 매크로가 끝난 뒤에도 브라우저 창을 남깁니다(Chrome, Edge). |
| `EnableBiDi()` **3.2** | WebDriver BiDi 연결을 요청합니다. `Start` 후 `WebSocketUrl`로 주소를 읽습니다. |
| `WebSocketUrl` **3.2** | BiDi 연결 주소입니다. `EnableBiDi`를 쓰지 않았으면 빈 문자열입니다. |
| `AddArgument(argument)` | 브라우저 실행 인수를 추가합니다. 예: `"--window-size=1280,800"` |
| `AddExtension(extensionPath)` | 확장 프로그램(.crx, .xpi)을 설치합니다. |
| `SetPreference(key, value)` | 브라우저 설정값을 지정합니다. 예: `"intl.accept_languages", "ko"` |
| `SetCapability(key, value)` | WebDriver capability를 직접 지정합니다. |
| `SetProfile(nameOrDirectory, [persistant])` | 사용자 프로필 폴더를 지정합니다. 로그인 상태를 유지할 때 씁니다. |
| `SetBinary(path)` | 브라우저 실행 파일 경로를 지정합니다. |
| `Proxy` | 프록시 설정 ([18장](#18-proxy-point-size-application)) |

```vba
Dim driver As New Selenium.ChromeDriver
driver.SetHeadless
driver.SetUserAgent "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
driver.HideAutomationInfo
driver.Start
```

### 3.3 페이지 이동과 정보

| 명령 | 설명 |
|---|---|
| `Get(url, [timeout], [raise])` | 페이지를 엽니다. `BaseUrl`을 지정했다면 상대 주소도 쓸 수 있습니다. |
| `BaseUrl` | 기준 주소(읽기/쓰기) |
| `GoBack()`, `GoForward()`, `Refresh()` | 뒤로, 앞으로, 새로 고침 |
| `Url`, `Title` | 현재 주소와 제목 |
| `PageSource()` | 페이지 HTML |
| `PageSourceMatch(pattern, [group])` | HTML에서 정규식과 처음 맞는 부분 |
| `PageSourceMatches(pattern, [group])` | HTML에서 정규식과 맞는 모든 부분(List) |
| `StatusCode` **3.2** | HTTP 상태 코드(200, 404 등). 알 수 없으면 0입니다. |
| `UserAgent` **3.2** | 브라우저의 User-Agent |
| `OpenHtml(html)` **3.2** | HTML 문자열을 브라우저에 엽니다. |
| `OpenFile(path)` **3.2** | 로컬 파일을 브라우저에 엽니다. |
| `PageToHtmlDocument()` **3.2** | 현재 페이지를 HTML 문서 객체(MSHTML)로 돌려줍니다. VBA에서 `getElementsByTagName` 등으로 바로 파싱할 수 있습니다. |
| `PageToJson()` **3.2** | JSON을 보여 주는 페이지(API 주소 등)를 Dictionary/List로 변환합니다. |
| `PageToXmlDocument()` **3.2** | 현재 페이지를 XML 문서 객체(MSXML2.DOMDocument.6.0)로 돌려줍니다. |

```vba
driver.Get "https://api.example.com/items"
Set data = driver.PageToJson
Debug.Print data("items")(1)("name")

Set doc = driver.PageToHtmlDocument
Debug.Print doc.getElementsByTagName("a").Length
```

### 3.4 요소 찾기

결과가 하나인 명령은 `WebElement`를, 여러 개인 명령은 `WebElements`를 돌려줍니다.

| 명령 | 찾는 기준 |
|---|---|
| `FindElementByCss(css, [timeout], [raise])` | CSS 선택자. 예: `"div.item > a"` |
| `FindElementByXPath(xpath, ...)` | XPath. 예: `"//button[text()='확인']"` |
| `FindElementById(id, ...)` | id 속성 |
| `FindElementByName(name, ...)` | name 속성 |
| `FindElementByClass(class, ...)` | class 이름 |
| `FindElementByTag(tag, ...)` | 태그 이름. 예: `"table"` |
| `FindElementByLinkText(text, ...)` | 링크 글자 전체 |
| `FindElementByPartialLinkText(text, ...)` | 링크 글자 일부 |
| `FindElement(by, ...)` | By 객체 ([7장](#7-by-찾는-방법)) |
| `FindElementBy(strategy, value, ...)` | 찾는 방식 번호와 값 |
| `FindElementsByCss(css, [minimum], [timeout])` 등 | 위와 같은 기준으로 **모두** 찾습니다. `minimum`개 이상 찾을 때까지 `timeout`만큼 기다립니다. |
| `IsElementPresent(by, [timeout])` | 요소가 있으면 True |
| `WaitNotElement(by, [timeout])` | 요소가 사라질 때까지 기다립니다. |
| `ActiveElement()` | 현재 포커스가 있는 요소 |

```vba
Set rows = driver.FindElementsByCss("table#list tr", minimum:=1, timeout:=5000)
For Each row In rows
    Debug.Print row.Text
Next
```

### 3.5 대기

| 명령 | 설명 |
|---|---|
| `Wait(timems)` | 지정한 시간만큼 멈춥니다. |
| `WaitForNetworkIdle([idleTime], [timeout])` **3.2** | 페이지 로딩이 끝나고, 네트워크 요청(fetch/XHR)이 `idleTime`(기본 500ms) 동안 없을 때까지 기다립니다. 기본 최대 30초입니다. |
| `WaitForDomIdle([idleTime], [timeout])` **3.2** | 화면(DOM) 변화가 `idleTime` 동안 없을 때까지 기다립니다. |
| `WaitForJQueryIdle([timeout])` **3.2** | jQuery 요청이 모두 끝날 때까지 기다립니다. jQuery가 없으면 바로 돌아옵니다. |
| `WaitForScript(script, [arguments], [timeout])` | JavaScript가 참(또는 Null이 아닌 값)을 돌려줄 때까지 기다립니다. |
| `Until(procedure, [argument], [timeout])` | 프로시저가 True를 돌려줄 때까지 반복합니다. |

```vba
driver.Get "https://example.com/search?q=abc"
driver.WaitForNetworkIdle          ' 검색 결과를 불러오는 요청이 끝날 때까지
```

### 3.6 스크롤 **3.2**

| 명령 | 설명 |
|---|---|
| `ScrollToTop()`, `ScrollToBottom()` | 맨 위로, 맨 아래로 |
| `ScrollTo(x, y)` | 지정 위치(픽셀)로 |
| `ScrollBy(x, y)` | 지정 거리만큼. 예: `ScrollBy 0, 500` |
| `ScrollUntilEnd([maxScrolls], [pause])` | 무한 스크롤 페이지를 더 불러올 내용이 없을 때까지 내립니다. 내용을 더 불러온 횟수를 돌려줍니다. `maxScrolls`는 최대 횟수(기본 100), `pause`는 매번 기다리는 시간(기본 1000ms)입니다. |
| `ScrollHeight`, `ScrollWidth` | 페이지 전체 높이와 너비(픽셀) |

```vba
n = driver.ScrollUntilEnd(maxScrolls:=30)
Debug.Print n & "번 더 불러옴, 항목 " & driver.FindElementsByCss(".item").Count & "개"
```

### 3.7 창, 탭, 프레임

| 명령 | 설명 |
|---|---|
| `Window` | 현재 창 ([8장](#8-window-창)) |
| `Windows` | 모든 창 목록(List) |
| `NewTab([url])` **3.2** | 새 탭을 열고 전환합니다. 주소를 주면 그 페이지도 엽니다. |
| `NewWindow([url])` **3.2** | 새 창을 열고 전환합니다. |
| `SwitchToNewWindow([timeout], [raise])` **3.2** | 클릭 등으로 **새로 열린** 창으로 전환합니다. 기준은 `Start`나 마지막 창 전환 시점입니다. |
| `SwitchToWindowByUrl(urlPattern, [timeout], [raise])` **3.2** | 주소가 패턴과 맞는 창으로 전환합니다. `*`와 `?`를 쓸 수 있습니다. 예: `"*/detail*"` |
| `SwitchToWindowByTitle(title, ...)` | 제목으로 창을 전환합니다. |
| `SwitchToWindowByName(name, ...)` | 창 이름으로 전환합니다. |
| `SwitchToNextWindow([timeout], [raise])` | 다음 창으로 전환합니다. |
| `SwitchToPreviousWindow()` | 직전 창으로 돌아갑니다. |
| `SwitchToFrame(identifier, [timeout], [raise])` | 프레임으로 들어갑니다. 번호, 이름, id, WebElement를 쓸 수 있습니다. |
| `SwitchToParentFrame()` | 상위 프레임으로 나옵니다. |
| `SwitchToDefaultContent()` | 프레임에서 완전히 나와 페이지 본문으로 돌아갑니다. |
| `CurrentFrameName` **3.2** | 현재 프레임 이름. 프레임 밖이면 빈 문자열입니다. |
| `SwitchToAlert([timeout], [raise])` | 알림창으로 전환합니다 ([11장](#11-alert-알림창)). |

```vba
driver.FindElementByLinkText("상세 보기").Click
driver.SwitchToNewWindow
Debug.Print driver.Title
driver.Close
driver.SwitchToPreviousWindow
```

### 3.8 다운로드와 PDF

| 명령 | 설명 |
|---|---|
| `SetDownloadFolder(folder)` **3.1** | 다운로드 폴더를 지정합니다. 저장 위치를 묻지 않고, PDF도 열지 않고 내려받습니다. Chrome과 Edge는 `Start` 전후 언제든, Firefox는 `Start` 전에만 쓸 수 있습니다. |
| `WaitForDownload([fileName], [timeout])` **3.1** | 다운로드가 끝날 때까지 기다리고 파일 경로를 돌려줍니다. `fileName`에 `"*.xlsx"`처럼 와일드카드를 쓸 수 있고, 기본 대기 시간은 60초입니다. |
| `DownloadResource(url, [path], [timeout])` **3.2** | 주소의 파일을 **로그인 상태(쿠키) 그대로** 내려받아 저장하고 경로를 돌려줍니다. `path`에 폴더를 주면 서버가 알려 준 파일 이름으로 저장합니다. |
| `PrintToPDF(filePath, [landscape], [scale], [background], [pageRanges], [paperSize], [margin], [shrinkToFit], [displayHeaderFooter], [headerTemplate], [footerTemplate])` **3.1/3.2** | 페이지를 PDF로 저장하고 경로를 돌려줍니다. 인수는 아래 표를 보세요. |

`PrintToPDF` 인수

| 인수 | 설명 |
|---|---|
| `filePath` | `{yyyyMMdd}`처럼 쓰면 날짜로 바뀝니다. |
| `landscape` | True면 가로 방향 |
| `scale` | 배율 0.1~2 (기본 1) |
| `background` | 배경 인쇄 여부 (기본 True) |
| `pageRanges` | 페이지 범위. 예: `"1-3,5"` |
| `paperSize` | `"A3"`, `"A4"`, `"A5"`, `"B4"`, `"B5"`, `"Letter"`, `"Legal"`, `"Tabloid"`, 또는 cm 단위의 `"21x29.7"` |
| `margin` | 상하좌우 여백(cm) |
| `shrinkToFit` | 페이지 너비에 맞추기 (기본 True) |
| `displayHeaderFooter` | 날짜, 제목, 주소, 쪽 번호를 머리글과 바닥글에 넣습니다(Chrome, Edge). |
| `headerTemplate`, `footerTemplate` | 머리글과 바닥글 HTML. `date`, `title`, `url`, `pageNumber`, `totalPages` 클래스를 쓸 수 있습니다(Chrome, Edge). |

```vba
driver.SetDownloadFolder "C:\Temp\보고서"
driver.FindElementById("excelDown").Click
path = driver.WaitForDownload("*.xlsx")
Workbooks.Open path

driver.PrintToPDF "C:\Temp\page_{yyyyMMdd}.pdf", paperSize:="A4", margin:=1, displayHeaderFooter:=True
```

### 3.9 JavaScript와 개발자 도구

| 명령 | 설명 |
|---|---|
| `ExecuteScript(script, [arguments])` | JavaScript를 실행하고 `return` 값을 돌려줍니다. 인수는 `arguments[0]`처럼 받습니다. |
| `ExecuteAsyncScript(script, [arguments], [timeout])` | 비동기 JavaScript를 실행합니다. `callback(값)`으로 결과를 돌려줍니다. |
| `ExecuteCDP(command, [parameters])` **3.2** | 크롬 개발자 도구(CDP) 명령을 실행합니다(Chrome, Edge). 매개변수는 Dictionary나 JSON 문자열로 넘깁니다. |
| `Send(method, relativeUri, ...)` | WebDriver 명령을 직접 보냅니다(고급). |

```vba
Debug.Print driver.ExecuteScript("return document.querySelectorAll('a').length")
driver.ExecuteCDP "Network.clearBrowserCache"
driver.ExecuteCDP "Emulation.setTimezoneOverride", "{""timezoneId"":""Asia/Seoul""}"
```

### 3.10 기타

| 명령 | 설명 |
|---|---|
| `TakeScreenshot([delayms])` | 화면을 캡처해 Image로 돌려줍니다. 예: `driver.TakeScreenshot.SaveAs "C:\a.png"` |
| `SendKeys(keysOrModifier, [keys])` | 현재 포커스가 있는 요소에 키를 입력합니다. |
| `Manage` | 쿠키, 저장소, 위치 ([9장](#9-manage-cookies-cookie-storage-쿠키와-저장소)) |
| `Timeouts` | 대기 시간 설정 ([10장](#10-timeouts-대기-시간)) |
| `Actions`, `Mouse`, `Keyboard`, `Keys` | 마우스와 키보드 ([12장](#12-actions-mouse-keyboard-keys-마우스와-키보드)) |
| `SetClipBoard(text)`, `GetClipBoard()` | 클립보드에 쓰기, 클립보드 읽기 |

---

## 4. WebElement: 요소

`FindElementBy...`로 찾은 요소입니다. 요소 안에서 다시 `FindElementBy...`로 찾을 수 있습니다([3.4](#34-요소-찾기)와 같은 명령).

### 4.1 읽기

| 명령 | 설명 |
|---|---|
| `Text()` | 화면에 보이는 글자 |
| `Value()` | 입력란의 값 |
| `Attribute(name)` | 속성이나 프로퍼티 값. 예: `Attribute("href")` |
| `Property(name)` **3.2** | JavaScript 프로퍼티만 읽습니다. 예: 체크 상태는 `Property("checked")` |
| `DomAttribute(name)` **3.2** | HTML에 적힌 속성만 읽습니다. 없으면 Nothing입니다. |
| `HasAttribute(name)` **3.2** | 속성이 있으면 True |
| `CssValue(propertyName)` | CSS 값. 예: `CssValue("color")` |
| `TagName` | 태그 이름 |
| `TextMatch(pattern)`, `TextMatches(pattern)` | 글자에서 정규식과 맞는 부분(처음 하나, 전부) |
| `TextAsNumber([decimalCharacter], [errorValue])` | 글자를 숫자로 변환합니다. 예: "1,234원" → 1234 |
| `IsDisplayed`, `IsEnabled`, `IsSelected`, `IsPresent` | 보이는지, 사용 가능한지, 선택됐는지, 아직 페이지에 있는지 |
| `HasFocus` **3.2** | 포커스가 있으면 True |
| `AriaLabel`, `AriaRole` **3.2** | 화면 읽기 프로그램이 읽는 이름과 역할(button, link 등) |
| `Location()`, `LocationInView()`, `Size()` | 위치와 크기 |
| `Parent` **3.2** | 부모 요소. 최상위면 Nothing입니다. |
| `Children()` **3.2** | 자식 요소 목록 |
| `ShadowRoot([timeout], [raise])` **3.1** | Shadow DOM 안쪽 ([6장](#6-shadowroot-shadow-dom)) |
| `Equals(obj)` | 같은 요소면 True |

### 4.2 조작

| 명령 | 설명 |
|---|---|
| `Click([modifierKeys])` | 클릭합니다. 예: `Click Keys.Control`은 Ctrl+클릭 |
| `ClickDouble()`, `ClickContext()` | 더블클릭, 오른쪽 클릭 |
| `ClickByOffset(x, y)` | 요소 왼쪽 위에서 x, y만큼 떨어진 곳을 클릭합니다. |
| `ClickAndHold()`, `ReleaseMouse()` | 누르고 있기, 놓기 |
| `SendKeys(keysOrModifier, [keys])` | 글자를 입력합니다. 예: `SendKeys "abc" & Keys.Enter` |
| `HoldKeys(modifierKeys)`, `ReleaseKeys(modifierKeys)` | Ctrl 같은 키를 누르고 있기, 놓기 |
| `Clear()` | 입력란을 비웁니다. |
| `Submit()` | 폼을 제출합니다. |
| `Focus()` **3.2** | 포커스를 줍니다. |
| `DragAndDropToWebElement(target)`, `DragAndDropToOffset(x, y)` | 다른 요소로, 또는 지정 거리만큼 끌어다 놓습니다. |
| `ScrollIntoView([alignTop])` | 요소가 보이도록 스크롤합니다. |
| `SetAttribute(name, value)` **3.2** | HTML 속성을 바꿉니다. 예: `SetAttribute "readonly", ""` |
| `RemoveAttribute(name)` **3.2** | HTML 속성을 지웁니다. 예: `RemoveAttribute "disabled"` |
| `Highlight([color])`, `Unhighlight()` **3.2** | 테두리 강조 표시(기본 빨강)와 해제. 디버깅할 때 유용합니다. |
| `UploadFile(path1, [path2], ...)` **3.2** | 파일 선택 상자(`input type="file"`)에 파일을 지정합니다. 여러 개나 배열도 됩니다. |
| `DownloadResource([path], [timeout])` **3.2** | 이미지(src)나 링크(href)의 파일을 로그인 상태 그대로 저장하고 경로를 돌려줍니다. |
| `TakeScreenshot([delayms])` | 요소만 캡처합니다. |
| `ExecuteScript(script, [arguments])` | 요소를 `this`로 하여 JavaScript를 실행합니다. 예: `ExecuteScript("return this.innerHTML")` |
| `ExecuteAsyncScript(...)`, `WaitForScript(...)` | 비동기 실행, 조건 대기 |

```vba
Set fileInput = driver.FindElementByCss("input[type=file]")
fileInput.UploadFile "C:\첨부\a.pdf", "C:\첨부\b.pdf"

path = driver.FindElementByCss("img.chart").DownloadResource("C:\Temp\")
```

### 4.3 요소 안쪽 스크롤 **3.2**

스크롤 막대가 있는 요소(목록 상자 등)에 씁니다. 사용법은 [3.6](#36-스크롤-32)과 같습니다.

`ScrollToTop()`, `ScrollToBottom()`, `ScrollTo(x, y)`, `ScrollBy(x, y)`, `ScrollUntilEnd([maxScrolls], [pause])`, `ScrollHeight`, `ScrollWidth`

### 4.4 대기

모두 기다린 뒤 요소 자신을 돌려주므로 이어서 쓸 수 있습니다.

| 명령 | 기다리는 조건 |
|---|---|
| `WaitDisplayed([displayed], [timeout])` | 보일 때까지. False면 숨겨질 때까지 |
| `WaitEnabled([enabled], [timeout])` | 사용 가능해질 때까지 |
| `WaitSelection([selected], [timeout])` | 선택될 때까지 |
| `WaitText(pattern, [timeout])`, `WaitNotText(...)` | 글자가 정규식과 맞을 때까지, 달라질 때까지 |
| `WaitAttribute(attribute, pattern, ...)`, `WaitNotAttribute(...)` | 속성 값 |
| `WaitCssValue(property, value, ...)`, `WaitNotCssValue(...)` | CSS 값 |
| `WaitRemoval([timeout])` | 요소가 페이지에서 사라질 때까지 |
| `WaitForDomIdle([idleTime], [timeout])` **3.2** | 요소 안쪽의 변화가 멈출 때까지 |
| `Until(procedure, [argument], [timeout])` | 프로시저가 True를 돌려줄 때까지 |

```vba
driver.FindElementById("result").WaitText "완료", 10000
```

### 4.5 변환

| 명령 | 설명 |
|---|---|
| `AsSelect()` | 선택 상자로 다룹니다 ([13장](#13-selectelement-tableelement-선택-상자와-표)). |
| `AsTable()` | 표로 다룹니다 ([13장](#13-selectelement-tableelement-선택-상자와-표)). |

---

## 5. WebElements: 요소 목록

`FindElementsBy...`의 결과입니다. `For Each`로 돌 수 있습니다.

| 명령 | 설명 |
|---|---|
| `Item(n)` 또는 `(n)` | n번째 요소(1부터) |
| `Count` | 개수 |
| `First()`, `Last()` | 처음, 마지막 요소 |
| `Text([offsetStart], [offsetEnd], [trim])` | 각 요소의 글자 목록(List) |
| `Attribute(name, [withAttributeOnly])` | 각 요소의 속성 목록 |
| `Values([defaultValue], ...)` | 각 요소의 글자를 숫자로 바꾼 목록 |
| `ExecuteScript(script, [includeNullResults])` | 요소마다 JavaScript를 실행한 결과 목록. 예: `"return element.href"` |
| `ExecuteAsyncScript(script, [timeout])` | 요소마다 비동기 실행 |
| `Highlight([color])`, `Unhighlight()` **3.2** | 모두 강조 표시, 해제 |

```vba
Set names = driver.FindElementsByCss("td.name").Text
names.ToExcel Range("A1")          ' 엑셀에 세로로 붙여 넣기
```

---

## 6. ShadowRoot: Shadow DOM **3.1**

웹 컴포넌트처럼 Shadow DOM 안에 숨겨진 요소를 찾습니다.

```vba
Set root = driver.FindElementByCss("my-component").ShadowRoot
root.FindElementByCss("button.save").Click
```

- 사용할 수 있는 명령: `FindElementBy...`, `FindElementsBy...`, `IsElementPresent`, `WaitNotElement`
- Chrome과 Edge는 Shadow DOM 안에서 **XPath와 링크 텍스트 검색을 지원하지 않습니다.** CSS, id, name, class, 태그로 찾으세요.
- `ShadowRoot(timeout, raise)`: Shadow DOM이 늦게 붙는 페이지에서는 기다립니다. 없으면 `NoSuchShadowRootError` 오류가 납니다. `raise:=False`면 Nothing을 돌려줍니다.

---

## 7. By: 찾는 방법

`FindElement`, `FindElements`, `IsElementPresent`, `WaitNotElement`에 넘깁니다.

| 명령 | 설명 |
|---|---|
| `Css(selector)`, `XPath(xpath)`, `Id(id)`, `Name(name)`, `Class(className)`, `Tag(tagName)`, `LinkText(text)`, `PartialLinkText(text)` | 찾는 기준 |
| `Any(by1, by2, [by3]...[by6])` | 여러 기준 중 하나라도 맞으면 찾습니다. |

```vba
Dim by As New Selenium.By
Set el = driver.FindElement(by.Any(by.Id("login"), by.Css(".btn-login")))
If driver.IsElementPresent(by.Css(".error"), 1000) Then MsgBox "오류 표시됨"
```

---

## 8. Window: 창

`driver.Window`(현재 창), `driver.Windows(n)`, `NewTab` 등의 결과입니다.

| 명령 | 설명 |
|---|---|
| `Activate()` | 이 창으로 전환합니다. |
| `Title` | 제목 |
| `Url` **3.2** | 주소 |
| `Handle` **3.2** | 창 식별자 |
| `Maximize()`, `Minimize()` **3.2**, `FullScreen()` | 최대화, 최소화, 전체 화면 |
| `Size()`, `SetSize(width, height)` | 창 크기 |
| `Position()`, `SetPosition(x, y)` | 창 위치 |
| `Close()` | 창을 닫습니다. |
| `SwitchToNextWindow(...)`, `SwitchToPreviousWindow()`, `SwitchToWindowByName(...)`, `SwitchToWindowByTitle(...)` | 다른 창으로 전환 |

---

## 9. Manage, Cookies, Cookie, Storage: 쿠키와 저장소

### Manage (`driver.Manage`)

| 명령 | 설명 |
|---|---|
| `Cookies` | 현재 사이트의 쿠키 목록 |
| `AddCookie(name, value, [domain], [path], [expiry], [secure], [httpOnly], [sameSite])` | 쿠키를 추가합니다. `sameSite`(**3.2**)에는 `"Strict"`, `"Lax"`, `"None"`을 씁니다. |
| `FindCookieByName(namePattern)` | 이름이 정규식과 맞는 쿠키 |
| `DeleteCookieByName(name)`, `DeleteAllCookies()` | 쿠키 삭제 |
| `StorageLocal`, `StorageSession` | localStorage, sessionStorage |
| `SetLocation(latitude, longitude, altitude)` | 페이지에 알려 줄 위치를 지정합니다(Chrome, Edge). 3.2부터 최신 드라이버에서 다시 동작합니다. |
| `Location()` | 페이지가 보는 현재 위치(Dictionary: latitude, longitude, altitude, accuracy) |
| `Capabilities` | 세션 정보(브라우저 이름, 버전 등) |

### Cookies

| 명령 | 설명 |
|---|---|
| `Item(이름 또는 번호)`, `Count` | 쿠키 하나, 개수 |
| `SaveToFile(filePath)` **3.2** | 쿠키를 JSON 파일로 저장합니다. |
| `LoadFromFile(filePath)` **3.2** | 저장한 쿠키를 다시 넣고, 넣은 개수를 돌려줍니다. **같은 사이트의 페이지를 먼저 열어야 합니다.** |

### Cookie

`Name`, `Value`, `Domain`, `Path`, `Expiry`, `Secure`, `HttpOnly` **3.2**, `SameSite` **3.2**, `Delete()`

```vba
' 처음 한 번: 직접 로그인한 뒤 저장
driver.Manage.Cookies.SaveToFile "C:\Temp\login.json"

' 다음부터: 사이트를 열고 쿠키를 넣은 뒤 새로 고침
driver.Get "https://example.com"
driver.Manage.Cookies.LoadFromFile "C:\Temp\login.json"
driver.Refresh
```

### Storage (`StorageLocal`, `StorageSession`)

`Item(key)` 읽기/쓰기, `Count()`, `Keys()`, `Remove(key)`, `Clear()`

---

## 10. Timeouts: 대기 시간

`driver.Timeouts`의 값을 밀리초 단위로 읽고 씁니다.

| 명령 | 설명 (기본값) |
|---|---|
| `ImplicitWait` | 요소를 찾을 때 기다리는 시간 (3000) |
| `PageLoad` | 페이지 로딩을 기다리는 시간 (60000) |
| `Script` | 비동기 JavaScript를 기다리는 시간 (15000) |
| `Server` | 드라이버 응답을 기다리는 시간 (90000) |

```vba
driver.Timeouts.ImplicitWait = 10000
```

---

## 11. Alert: 알림창

```vba
Set dlg = driver.SwitchToAlert(3000, False)
If Not dlg Is Nothing Then
    Debug.Print dlg.Text
    dlg.Accept
End If
```

| 명령 | 설명 |
|---|---|
| `Text` | 알림창 내용 |
| `Accept()`, `Dismiss()` | 확인, 취소 |
| `SendKeys(keysToSend)` | 입력창(prompt)에 글자를 입력합니다. |
| `SetCredentials(user, password)` | 로그인 창에 아이디와 비밀번호를 입력합니다. |

---

## 12. Actions, Mouse, Keyboard, Keys: 마우스와 키보드

### Actions (`driver.Actions`)

여러 동작을 이어 붙인 뒤 `Perform`으로 한 번에 실행합니다.

```vba
driver.Actions.MoveToElement(menu).Wait(300).Click(subMenu).Perform
driver.Actions.KeyDown(Keys.Control).Click(item1).Click(item2).KeyUp(Keys.Control).Perform
```

| 명령 | 설명 |
|---|---|
| `Click([element])`, `ClickDouble([element])`, `ClickContext([element])` | 클릭, 더블클릭, 오른쪽 클릭 |
| `ClickAndHold([element])`, `Release([element])` | 누르고 있기, 놓기 |
| `MoveToElement(element)` | 요소 가운데로 마우스 이동 |
| `MoveByOffset(x, y)` | 현재 위치에서 상대 이동 |
| `MoveToPosition(x, y)` **3.2** | 화면(뷰포트)의 지정 좌표로 이동 |
| `DragAndDrop(source, target)` | 요소를 다른 요소로 끌어다 놓기 |
| `DragAndDropByOffset(source, x, y)` | 지정 거리만큼 끌기 |
| `DragAndDropToPosition(element, x, y)` **3.2** | 화면의 지정 좌표로 끌기 |
| `ScrollByWheel(deltaX, deltaY, [element])` **3.2** | 마우스 휠로 스크롤 |
| `ScrollToElement(element)` **3.2** | 요소가 보이도록 휠 스크롤 |
| `KeyDown(modifierKeys, [element])`, `KeyUp(modifierKeys)` | Ctrl, Shift, Alt 누르기, 놓기 |
| `SendKeys(keys, [element])` | 글자 입력 |
| `Wait(timems)` | 동작 사이에 쉬기 |
| `Perform()` | 쌓아 둔 동작을 실행 |

### Mouse, Keyboard

| 객체 | 명령 |
|---|---|
| `driver.Mouse` | `MoveTo(element, [x], [y])`, `Click([button])`, `ClickDouble()`, `ClickAndHold([button])`, `Release([button])` |
| `driver.Keyboard` | `SendKeys(...)`, `KeyDown(modifiers)`, `KeyUp(modifiers)` |

### Keys

특수 키 문자를 돌려줍니다. 글자와 이어 붙여 씁니다.

```vba
Dim Keys As New Selenium.Keys
el.SendKeys "검색어" & Keys.Enter
el.SendKeys Keys.Chord(Keys.Control, "a")            ' Ctrl+A (3.2)
el.SendKeys Keys.Repeat(Keys.Backspace, 5)           ' 백스페이스 5번 (3.2)
```

| 분류 | 키 |
|---|---|
| 조합 | `Control`, `Shift`, `Alt`, `Meta`, `Command`, `LeftControl`, `LeftShift`, `LeftAlt` |
| 이동 | `Up`, `Down`, `Left`, `Right`, `ArrowUp`, `ArrowDown`, `ArrowLeft`, `ArrowRight`, `Home`, `End`, `PageUp`, `PageDown`, `Tab` |
| 편집 | `Enter`, `Return`, `Backspace`, `Delete`, `Insert`, `Escape`, `Space`, `Clear` |
| 기능 | `F1` ~ `F12`, `Help`, `Pause`, `Cancel`, `Null` |
| 숫자 패드 | `NumPad0` ~ `NumPad9`, `Add`, `Subtract`, `Multiply`, `Divide`, `Decimal`, `Separator`, `Equal`, `Semicolon` |
| 도우미 | `Repeat(keys, count)` **3.2**: 반복, `Chord(key1, [key2], [key3], [key4])` **3.2**: 함께 누르고 조합 키 놓기 |

---

## 13. SelectElement, TableElement: 선택 상자와 표

### SelectElement (`element.AsSelect`)

```vba
Set sel = driver.FindElementById("city").AsSelect
sel.SelectByText "서울"
Debug.Print sel.SelectedOption.Text
```

| 명령 | 설명 |
|---|---|
| `SelectByText(text)`, `SelectByValue(value)`, `SelectByIndex(index)` | 선택 |
| `DeselectByText(...)`, `DeselectByValue(...)`, `DeselectByIndex(...)`, `DeselectAll()` | 선택 해제(다중 선택 상자) |
| `Options`, `SelectedOption`, `AllSelectedOptions` | 전체 항목, 선택된 항목 |
| `IsMultiple` | 다중 선택 상자면 True |

### TableElement (`element.AsTable`)

| 명령 | 설명 |
|---|---|
| `Data([firstRowsToSkip], [lastRowsToSkip], [map])` | 표 내용을 2차원 배열로 돌려줍니다. |
| `ToExcel([target], [clear], [firstRowsToSkip], [lastRowsToSkip], [map])` | 표를 엑셀에 붙여 넣습니다. `target`에는 주소, Range, 시트를 쓸 수 있습니다. |

```vba
driver.FindElementByCss("table.result").AsTable.ToExcel Sheet1.Range("A1"), clear:=True
```

---

## 14. Utils, Waiter: 도구

### Utils (`New Selenium.Utils`)

| 명령 | 설명 |
|---|---|
| `ParseJson(json)` **3.2** | JSON 문자열을 Dictionary/List로 변환합니다. |
| `ToJson(value)` **3.2** | Dictionary, List, 배열, 문자열, 숫자를 JSON 문자열로 변환합니다. |
| `SaveBase64ToFile(base64, filePath)` **3.2** | Base64 문자열(`data:` 주소도 가능)을 파일로 저장하고 경로를 돌려줍니다. |
| `FileToBase64(filePath)` **3.2** | 파일을 Base64 문자열로 읽습니다. |
| `IsMatch(input, pattern)` | 정규식과 맞으면 True |
| `Match(input, pattern, [group])`, `Matches(...)` | 정규식과 맞는 부분(처음 하나, 전부) |
| `Replace(input, pattern, replacement)` | 정규식으로 바꾸기 |
| `TakeScreenShot()` | 바탕 화면 전체 캡처 |
| `SendKeysNat(keys)` | Windows SendKeys로 현재 프로그램에 키 입력 |

```vba
Dim u As New Selenium.Utils
Set obj = u.ParseJson("{""name"":""홍길동"",""tags"":[""a"",""b""]}")
Debug.Print obj("name"), obj("tags")(2)
u.SaveBase64ToFile driver.ExecuteScript("return document.querySelector('canvas').toDataURL()"), "C:\Temp\chart.png"
```

### Waiter (`New Selenium.Waiter`)

| 명령 | 설명 |
|---|---|
| `Wait(timems)` | 멈추기 |
| `Until(function, [argument], [timeout], [timeoutMessage])` | 함수가 True를 돌려줄 때까지 반복 |
| `Not(result, [timeout], [timeoutMessage])` | `Do While waiter.Not(조건)` 형태의 반복문에 씁니다. 시간이 지나면 오류를 냅니다. |
| `WaitForFile(path, [timeout])` | 파일이 생기고 쓰기가 끝날 때까지 기다립니다. |
| `Timeout` | 기본 대기 시간 (30000) |

---

## 15. Image, PdfFile: 이미지와 PDF

### Image (`TakeScreenshot`의 결과, 또는 `New Selenium.Image`)

| 명령 | 설명 |
|---|---|
| `SaveAs(filePath)` | png, bmp, gif, jpg로 저장합니다. 경로의 `{yyyyMMdd-HHmmss}`는 날짜로 바뀝니다. |
| `ToExcel([target])` | 엑셀 시트에 붙여 넣습니다. |
| `Copy()` | 클립보드로 복사합니다. |
| `Load(filePath)` | 이미지 파일을 불러옵니다. |
| `Resize([width], [height])` | 크기 변경 |
| `CompareTo(image, ...)`, `DiffCount` | 다른 이미지와 비교하고, 다른 픽셀 수를 셉니다. |
| `Width`, `Height`, `CRC` | 크기, 해시값 |
| `GetPicture()` | StdPicture 객체(이미지 컨트롤 등에 사용) |

### PdfFile (`New Selenium.PdfFile`)

스크린샷과 글을 모아 PDF 보고서를 만듭니다. 페이지를 그대로 PDF로 저장하려면 `driver.PrintToPDF`가 더 간단합니다.

`AddTitle`, `AddText`, `AddTextCenter`, `AddImage(image)`, `AddLink`, `AddBookmark`, `AddSpace`, `AddPage`, `SetMargins`, `SetPageSize`, `SaveAs(filepath)`

---

## 16. List, Dictionary, Table: 데이터

### List

| 명령 | 설명 |
|---|---|
| `Item(n)` / `(n)`, `Get(n)`, `Set(n, value)` | n번째 값(1부터) |
| `Add(value)`, `AddRange(array)`, `Insert(n, value)`, `RemoveAt(n)` | 추가, 삽입, 삭제 |
| `Count`, `First()`, `Last()`, `Contains(value)` | 개수, 처음, 마지막, 포함 여부 |
| `Sort()`, `Distinct([ignoreCase])` | 정렬, 중복 제거 |
| `Values()` | 배열로 변환 |
| `ToExcel([target], [title], [clear])` | 엑셀에 세로로 붙여 넣기 |

### Dictionary (`ExecuteScript`, `PageToJson` 등이 돌려주는 객체)

| 명령 | 설명 |
|---|---|
| `Item(key)` / `(key)` | 값 읽기/쓰기 |
| `Get(key, defaultValue)` | 키가 없으면 기본값 |
| `Add(key, value)`, `Set(key, value)`, `Clear()` | 추가, 설정, 비우기 |
| `ContainsKey(key)`, `Keys()`, `Values()`, `Count` | 키 확인, 키 배열, 값 배열, 개수 |
| `ToExcel([target], [clear])` | 키와 값을 엑셀에 붙여 넣기 |

### Table (`New Selenium.Table`)

엑셀 범위를 읽어 조건으로 거르는 데 씁니다.

`From(sourceExcel, [hasHeaders])`, `Where(expression, [sortBy])`, `Item(n)`, `Count`, `Values()`, `ToExcel([target], [clear])`

---

## 17. Assert, Verify: 검증

테스트 자동화에서 결과를 확인할 때 씁니다.

| 객체 | 동작 |
|---|---|
| `Assert` | 조건이 틀리면 **오류를 냅니다.** |
| `Verify` | 조건이 맞으면 `"OK"`, 틀리면 `"NOK 메시지"` 문자열을 **돌려줍니다.** |

명령(두 객체 공통): `True(value)`, `False(value)`, `Matches(pattern, input)`, `NotMatches(pattern, input)`, `Contains(value, input)`. `Assert`에는 `Fail([message])`도 있습니다.

```vba
Dim Assert As New Selenium.Assert
Assert.Contains "로그인 성공", driver.FindElementById("msg").Text
```

---

## 18. Proxy, Point, Size, Application

### Proxy (`driver.Proxy`, `Start` 전에 설정)

| 명령 | 설명 |
|---|---|
| `SetHttpProxy("host:port")`, `SetSSLProxy(...)`, `SetFTPProxy(...)` | 프로토콜별 프록시 |
| `SetSocksProxy(value, [username], [password])` | SOCKS 프록시 |
| `SetAutoConfigure(url)` | PAC 파일 주소 |
| `SetAutodetect()` | 자동 감지 |

### Point, Size

`Location()`, `Size()` 등의 결과입니다. `Point.X`, `Point.Y`, `Size.Width`, `Size.Height`

### Application (`New Selenium.Application`)

VBScript에서 여러 객체를 한 번에 만들 때 씁니다. 속성으로 `By`, `Keys`, `Utils`, `Waiter`, `Assert`, `Verify`가 있고, 메서드로 `ChromeDriver()`, `FirefoxDriver()`, `WebDriver()`, `List()`, `Dictionary()`, `Table()`, `PdfFile()`, `Point(x, y)`, `Size(w, h)`가 있습니다.

---

## 19. 현재 드라이버에서 동작하지 않는 명령

2.0.9와 호환되도록 남겨 두었지만, 최신 브라우저와 드라이버(W3C 방식)가 지원하지 않는 명령입니다.

| 명령 | 이유와 대안 |
|---|---|
| `TouchScreen`, `TouchActions` | 옛 모바일 전용 명령입니다. `Actions`를 쓰세요. |
| `CacheStatus()` | 브라우저에서 HTML5 애플리케이션 캐시가 제거되었습니다. |
| `IEDriver`, `OperaDriver`, `PhantomJSDriver` | 3.0에서 제거되었습니다. `Start`하면 오류가 납니다. Chrome이나 Edge를 쓰세요. |

---

## 오류가 날 때

| 오류 | 원인과 해결 |
|---|---|
| `NoSuchElementError` | 요소를 못 찾았습니다. 선택자를 확인하거나 `timeout`을 늘리세요. 프레임 안이면 `SwitchToFrame`, Shadow DOM 안이면 `ShadowRoot`를 쓰세요. |
| `ElementNotVisibleError` | 요소가 보이지 않거나 가려져 있습니다. `ScrollIntoView`나 `WaitDisplayed`를 먼저 쓰세요. |
| `ElementClickInterceptedError` | 다른 요소(팝업 등)가 가리고 있습니다. |
| `StaleElementReferenceError` | 페이지가 바뀌어 예전에 찾은 요소가 사라졌습니다. 다시 찾으세요. |
| `TimeoutError` | 기다리는 시간이 지났습니다. `Start`에서 "failed to open the listening port"라고 나오면 3.2.1 이상으로 업데이트하세요(Windows 업데이트 KB5127936 문제). |
| `Failed to get chromedriver.exe ...` | 드라이버를 내려받지 못했습니다. [README](../README.md#drivers)의 직접 넣는 방법을 보세요. |
| 런타임 오류 `8013101B` / `80131040` | 예전 버전 등록이 남아 있습니다. 엑셀을 닫고 설치 파일을 다시 실행하세요. |

진단 로그가 필요하면 `%LOCALAPPDATA%\SeleniumBasic\debug.log` 빈 파일을 만든 뒤 오류를 재현하세요.
