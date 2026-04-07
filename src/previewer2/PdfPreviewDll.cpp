/* Copyright 2022 the SumatraPDF project authors (see AUTHORS file).
   License: GPLv3 */

#include "utils/BaseUtil.h"
#include "utils/ScopedWin.h"
#include "utils/FileUtil.h"
#include "utils/WinUtil.h"

#include <Thumbcache.h>

#include "RegistryPreview.h"
#include "../PreviewPipe.h"

#include "utils/Log.h"

constexpr COLORREF kColWindowBg = RGB(0x99, 0x99, 0x99);
constexpr int kPreviewMargin = 2;
constexpr UINT kUwmPaintAgain = (WM_USER + 101);

class PageRenderer;

class PdfPreview : public IThumbnailProvider,
                   public IInitializeWithStream,
                   public IObjectWithSite,
                   public IPreviewHandler,
                   public IOleWindow {
  public:
    PdfPreview(long* plRefCount, PreviewFileType fileType) : m_fileType(fileType) {
        m_plModuleRef = plRefCount;
        InterlockedIncrement(m_plModuleRef);
    }

    ~PdfPreview() {
        Unload();
        InterlockedDecrement(m_plModuleRef);
    }

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) {
        static const QITAB qit[] = {QITABENT(PdfPreview, IInitializeWithStream),
                                    QITABENT(PdfPreview, IThumbnailProvider),
                                    QITABENT(PdfPreview, IObjectWithSite),
                                    QITABENT(PdfPreview, IPreviewHandler),
                                    QITABENT(PdfPreview, IOleWindow),
                                    {0}};
        return QISearch(this, qit, riid, ppv);
    }
    IFACEMETHODIMP_(ULONG) AddRef() { return InterlockedIncrement(&m_lRef); }
    IFACEMETHODIMP_(ULONG) Release() {
        long cRef = InterlockedDecrement(&m_lRef);
        if (cRef == 0) {
            delete this;
        }
        return cRef;
    }

    // IThumbnailProvider
    IFACEMETHODIMP GetThumbnail(uint cx, HBITMAP* phbmp, WTS_ALPHATYPE* pdwAlpha);

    // IInitializeWithStream
    IFACEMETHODIMP Initialize(IStream* pStm, __unused DWORD grfMode) {
        m_pStream = pStm;
        if (!m_pStream) {
            return E_INVALIDARG;
        }
        m_pStream->AddRef();
        return S_OK;
    };

    // IObjectWithSite
    IFACEMETHODIMP SetSite(IUnknown* punkSite) {
        m_site = nullptr;
        if (!punkSite) {
            return S_OK;
        }
        return punkSite->QueryInterface(&m_site);
    }
    IFACEMETHODIMP GetSite(REFIID riid, void** ppv) {
        if (m_site) {
            return m_site->QueryInterface(riid, ppv);
        }
        if (!ppv) {
            return E_INVALIDARG;
        }
        *ppv = nullptr;
        return E_FAIL;
    }

    // IPreviewHandler
    IFACEMETHODIMP SetWindow(HWND hwnd, const RECT* prc) {
        if (!hwnd || !prc) {
            return S_OK;
        }
        m_hwndParent = hwnd;
        return SetRect(prc);
    }
    IFACEMETHODIMP SetFocus() {
        if (!m_hwnd) {
            return S_FALSE;
        }
        ::SetFocus(m_hwnd);
        return S_OK;
    }
    IFACEMETHODIMP QueryFocus(HWND* phwnd) {
        if (!phwnd) {
            return E_INVALIDARG;
        }
        *phwnd = GetFocus();
        if (!*phwnd) {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        return S_OK;
    }
    IFACEMETHODIMP TranslateAccelerator(MSG* pmsg) {
        if (!m_site) {
            return S_FALSE;
        }
        ScopedComQIPtr<IPreviewHandlerFrame> frame(m_site);
        if (!frame) {
            return S_FALSE;
        }
        return frame->TranslateAccelerator(pmsg);
    }
    IFACEMETHODIMP SetRect(const RECT* prc) {
        if (!prc) {
            return E_INVALIDARG;
        }
        m_rcParent = ToRect(*prc);
        if (m_hwnd) {
            UINT flags = SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE;
            int x = m_rcParent.x;
            int y = m_rcParent.y;
            int dx = m_rcParent.dx;
            int dy = m_rcParent.dy;
            SetWindowPos(m_hwnd, nullptr, x, y, dx, dy, flags);
            InvalidateRect(m_hwnd, nullptr, TRUE);
            UpdateWindow(m_hwnd);
        }
        return S_OK;
    }
    IFACEMETHODIMP DoPreview();
    IFACEMETHODIMP Unload() {
        if (m_hwnd) {
            DestroyWindow(m_hwnd);
            m_hwnd = nullptr;
        }
        m_pStream = nullptr;
        if (pipeSession) {
            delete pipeSession;
            pipeSession = nullptr;
        }
        return S_OK;
    }

    // IOleWindow
    IFACEMETHODIMP GetWindow(HWND* phwnd) {
        if (!m_hwndParent || !phwnd) {
            return E_INVALIDARG;
        }
        *phwnd = m_hwndParent;
        return S_OK;
    }
    IFACEMETHODIMP ContextSensitiveHelp(__unused BOOL fEnterMode) { return E_NOTIMPL; }

    PreviewFileType GetFileType() { return m_fileType; }

    PageRenderer* renderer = nullptr;
    PreviewPipeSession* pipeSession = nullptr;

  private:
    long m_lRef = 1;
    long* m_plModuleRef = nullptr;
    ScopedComPtr<IStream> m_pStream;
    ScopedComPtr<IUnknown> m_site;
    HWND m_hwnd = nullptr;
    HWND m_hwndParent = nullptr;
    Rect m_rcParent;
    PreviewFileType m_fileType;

    HBITMAP GetThumbnailViaPipe(uint cx);
    bool InitPreviewSession();
};

HBITMAP PdfPreview::GetThumbnailViaPipe(uint cx) {
    logf("GetThumbnailViaPipe: cx=%d\n", cx);

    // Read stream data
    ByteSlice data = GetDataFromStream(m_pStream.Get(), nullptr);
    if (data.empty()) {
        logf("GetThumbnailViaPipe: failed to get data from stream\n");
        return nullptr;
    }

    logf("GetThumbnailViaPipe: read %d bytes from stream\n", (int)data.size());

    // Generate unique pipe name
    char* pipeName = GenerateUniquePipeName();
    if (!pipeName) {
        logf("GetThumbnailViaPipe: failed to generate pipe name\n");
        data.Free();
        return nullptr;
    }

    logf("GetThumbnailViaPipe: pipe name '%s'\n", pipeName);

    // Create named pipe (we are the server)
    HANDLE hPipe = CreatePreviewPipe(pipeName);
    if (hPipe == INVALID_HANDLE_VALUE) {
        logf("GetThumbnailViaPipe: failed to create pipe\n");
        str::Free(pipeName);
        data.Free();
        return nullptr;
    }

    // Launch SumatraPDF.exe
    HANDLE hProcess = LaunchSumatraForPreview(pipeName);
    if (!hProcess) {
        logf("GetThumbnailViaPipe: failed to launch SumatraPDF\n");
        CloseHandle(hPipe);
        str::Free(pipeName);
        data.Free();
        return nullptr;
    }

    str::Free(pipeName);

    HBITMAP result = nullptr;

    // Wait for client to connect
    if (!WaitForPipeConnection(hPipe)) {
        logf("GetThumbnailViaPipe: pipe connection failed\n");
        goto cleanup;
    }

    logf("GetThumbnailViaPipe: client connected\n");

    // Send request
    if (!SendPreviewRequest(hPipe, GetFileType(), cx, data)) {
        logf("GetThumbnailViaPipe: failed to send request\n");
        goto cleanup;
    }

    logf("GetThumbnailViaPipe: request sent\n");

    // Receive response
    result = ReceivePreviewResponse(hPipe);
    if (result) {
        logf("GetThumbnailViaPipe: received bitmap\n");
    } else {
        logf("GetThumbnailViaPipe: failed to receive response\n");
    }

cleanup:
    DisconnectNamedPipe(hPipe);
    CloseHandle(hPipe);

    // Terminate the process if still running
    TerminateProcess(hProcess, 0);
    CloseHandle(hProcess);

    data.Free();

    return result;
}

IFACEMETHODIMP PdfPreview::GetThumbnail(uint cx, HBITMAP* phbmp, WTS_ALPHATYPE* pdwAlpha) {
    logf("PdfPreview::GetThumbnail(cx=%d)\n", (int)cx);

    // Use pipe communication to SumatraPDF.exe for thumbnail generation
    HBITMAP hBitmap = GetThumbnailViaPipe(cx);
    if (!hBitmap) {
        logf("PdfPreview::GetThumbnail: GetThumbnailViaPipe failed\n");
        return E_FAIL;
    }

    // Verify bitmap before returning to Explorer
    BITMAP bm{};
    if (GetObject(hBitmap, sizeof(bm), &bm)) {
        logf("PdfPreview::GetThumbnail: returning HBITMAP %dx%d, bitsPixel=%d, bmHeight=%d (negative=top-down)\n",
             bm.bmWidth, bm.bmHeight, bm.bmBitsPixel, bm.bmHeight);
    }

    // Debug: save thumbnail to temp file for visual verification
    {
        TempStr tmpDir = GetTempDirTemp();
        TempStr debugPath = path::JoinTemp(tmpDir, "sumatrapdf-thumb-debug.bmp");
        // Write BMP file
        int absHeight = bm.bmHeight < 0 ? -bm.bmHeight : bm.bmHeight;
        u32 dataSize = (u32)(bm.bmWidth * absHeight * 4);
        BITMAPFILEHEADER bfh{};
        bfh.bfType = 0x4D42; // "BM"
        bfh.bfOffBits = sizeof(BITMAPFILEHEADER) + sizeof(BITMAPINFOHEADER);
        bfh.bfSize = bfh.bfOffBits + dataSize;
        BITMAPINFOHEADER bih{};
        bih.biSize = sizeof(bih);
        bih.biWidth = bm.bmWidth;
        bih.biHeight = bm.bmHeight; // preserve top-down/bottom-up
        bih.biPlanes = 1;
        bih.biBitCount = 32;
        bih.biCompression = BI_RGB;
        HANDLE hFile = CreateFileA(debugPath, GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (hFile != INVALID_HANDLE_VALUE) {
            DWORD written;
            WriteFile(hFile, &bfh, sizeof(bfh), &written, nullptr);
            WriteFile(hFile, &bih, sizeof(bih), &written, nullptr);
            WriteFile(hFile, bm.bmBits, dataSize, &written, nullptr);
            CloseHandle(hFile);
            logf("PdfPreview::GetThumbnail: debug bitmap saved to '%s'\n", debugPath);
        }
    }

    *phbmp = hBitmap;
    if (pdwAlpha) {
        *pdwAlpha = WTSAT_RGB;
        logf("PdfPreview::GetThumbnail: set alpha type to WTSAT_RGB (%d)\n", (int)WTSAT_RGB);
    } else {
        logf("PdfPreview::GetThumbnail: pdwAlpha is null\n");
    }
    logf("PdfPreview::GetThumbnail: returning S_OK\n");
    return S_OK;
}

// Initialize a pipe session for version 2 protocol (session-based preview)
bool PdfPreview::InitPreviewSession() {
    logf("InitPreviewSession\n");

    // Read stream data
    ByteSlice data = GetDataFromStream(m_pStream.Get(), nullptr);
    if (data.empty()) {
        logf("InitPreviewSession: failed to get data from stream\n");
        return false;
    }

    logf("InitPreviewSession: read %d bytes from stream\n", (int)data.size());

    // Generate unique pipe name
    char* pipeName = GenerateUniquePipeName();
    if (!pipeName) {
        logf("InitPreviewSession: failed to generate pipe name\n");
        data.Free();
        return false;
    }

    logf("InitPreviewSession: pipe name '%s'\n", pipeName);

    // Create named pipe (we are the server)
    HANDLE hPipe = CreatePreviewPipe(pipeName);
    if (hPipe == INVALID_HANDLE_VALUE) {
        logf("InitPreviewSession: failed to create pipe\n");
        str::Free(pipeName);
        data.Free();
        return false;
    }

    // Launch SumatraPDF.exe
    HANDLE hProcess = LaunchSumatraForPreview(pipeName);
    if (!hProcess) {
        logf("InitPreviewSession: failed to launch SumatraPDF\n");
        CloseHandle(hPipe);
        str::Free(pipeName);
        data.Free();
        return false;
    }

    str::Free(pipeName);

    // Wait for client to connect
    if (!WaitForPipeConnection(hPipe)) {
        logf("InitPreviewSession: pipe connection failed\n");
        CloseHandle(hPipe);
        TerminateProcess(hProcess, 0);
        CloseHandle(hProcess);
        data.Free();
        return false;
    }

    logf("InitPreviewSession: client connected\n");

    // Send Init command (version 2 protocol)
    DWORD bytesWritten = 0, bytesRead = 0;

    u32 magic = kPreviewRequestMagic;
    u32 version = kPreviewProtocolVersion2;
    u32 cmd = (u32)PreviewCmd::Init;
    u32 fileType = (u32)GetFileType();
    u32 dataSize = (u32)data.size();

    WriteFile(hPipe, &magic, 4, &bytesWritten, nullptr);
    WriteFile(hPipe, &version, 4, &bytesWritten, nullptr);
    WriteFile(hPipe, &cmd, 4, &bytesWritten, nullptr);
    WriteFile(hPipe, &fileType, 4, &bytesWritten, nullptr);
    WriteFile(hPipe, &dataSize, 4, &bytesWritten, nullptr);

    // Write file data
    DWORD totalWritten = 0;
    while (totalWritten < dataSize) {
        DWORD toWrite = dataSize - totalWritten;
        if (!WriteFile(hPipe, data.data() + totalWritten, toWrite, &bytesWritten, nullptr) || bytesWritten == 0) {
            logf("InitPreviewSession: failed to write file data\n");
            CloseHandle(hPipe);
            TerminateProcess(hProcess, 0);
            CloseHandle(hProcess);
            data.Free();
            return false;
        }
        totalWritten += bytesWritten;
    }

    FlushFileBuffers(hPipe);
    data.Free();

    // Read Init response: magic(4) + status(4) + pageCount(4)
    u32 respMagic = 0, status = 0, pageCount = 0;
    if (!ReadFile(hPipe, &respMagic, 4, &bytesRead, nullptr) || bytesRead != 4 || respMagic != kPreviewResponseMagic) {
        logf("InitPreviewSession: invalid response magic\n");
        CloseHandle(hPipe);
        TerminateProcess(hProcess, 0);
        CloseHandle(hProcess);
        return false;
    }
    if (!ReadFile(hPipe, &status, 4, &bytesRead, nullptr) || bytesRead != 4 || status != 0) {
        logf("InitPreviewSession: init failed with status %d\n", status);
        CloseHandle(hPipe);
        TerminateProcess(hProcess, 0);
        CloseHandle(hProcess);
        return false;
    }
    if (!ReadFile(hPipe, &pageCount, 4, &bytesRead, nullptr) || bytesRead != 4) {
        logf("InitPreviewSession: failed to read page count\n");
        CloseHandle(hPipe);
        TerminateProcess(hProcess, 0);
        CloseHandle(hProcess);
        return false;
    }

    logf("InitPreviewSession: success, pageCount=%d\n", pageCount);

    // Create pipe session
    pipeSession = new PreviewPipeSession();
    pipeSession->hPipe = hPipe;
    pipeSession->hProcess = hProcess;
    pipeSession->pageCount = (int)pageCount;

    return true;
}

// PageRenderer class using pipe session
class PageRenderer {
    PreviewPipeSession* session = nullptr;
    HWND hwnd = nullptr;

    int currPage = 0;
    HBITMAP currBmp = nullptr;
    Size currSize;

    int reqPage = 0;
    float reqZoom = 0.f;
    Size reqSize = {};
    bool reqAbort = false;

    CRITICAL_SECTION currAccess;
    HANDLE thread = nullptr;

  public:
    PageRenderer(PreviewPipeSession* session, HWND hwnd) {
        this->session = session;
        this->hwnd = hwnd;
        InitializeCriticalSection(&currAccess);
    }

    ~PageRenderer() {
        if (thread) {
            reqAbort = true;
            WaitForSingleObject(thread, INFINITE);
        }
        if (currBmp) {
            DeleteObject(currBmp);
        }
        DeleteCriticalSection(&currAccess);
    }

    RectF GetPageRect(int pageNo) {
        if (!session || !session->IsConnected()) {
            return RectF();
        }
        return session->GetPageBox(pageNo);
    }

    void Render(HDC hdc, Rect target, int pageNo, float zoom) {
        log("PageRenderer::Render()\n");

        ScopedCritSec scope(&currAccess);
        if (currBmp && currPage == pageNo && currSize == target.Size()) {
            // Blit cached bitmap
            HDC hdcMem = CreateCompatibleDC(hdc);
            HGDIOBJ oldBmp = SelectObject(hdcMem, currBmp);
            BitBlt(hdc, target.x, target.y, target.dx, target.dy, hdcMem, 0, 0, SRCCOPY);
            SelectObject(hdcMem, oldBmp);
            DeleteDC(hdcMem);
        } else if (!thread) {
            reqPage = pageNo;
            reqZoom = zoom;
            reqSize = target.Size();
            reqAbort = false;
            thread = CreateThread(nullptr, 0, RenderThread, this, 0, nullptr);
        } else if (reqPage != pageNo || reqSize != target.Size()) {
            reqAbort = true;
        }
    }

  protected:
    static DWORD WINAPI RenderThread(LPVOID data) {
        log("PageRenderer::RenderThread started\n");
        ScopedCom comScope;

        PageRenderer* pr = (PageRenderer*)data;

        if (!pr->session || !pr->session->IsConnected()) {
            return 0;
        }

        HBITMAP bmp = pr->session->RenderPage(pr->reqPage, pr->reqZoom, pr->reqSize.dx, pr->reqSize.dy);
        if (!bmp) {
            log("PageRenderer::RenderThread: RenderPage failed\n");
            ScopedCritSec scope(&pr->currAccess);
            HANDLE th = pr->thread;
            pr->thread = nullptr;
            CloseHandle(th);
            DestroyTempAllocator();
            return 0;
        }

        ScopedCritSec scope(&pr->currAccess);

        if (!pr->reqAbort) {
            if (pr->currBmp) {
                DeleteObject(pr->currBmp);
            }
            pr->currBmp = bmp;
            pr->currPage = pr->reqPage;
            pr->currSize = pr->reqSize;
        } else {
            DeleteObject(bmp);
        }

        HANDLE th = pr->thread;
        pr->thread = nullptr;
        PostMessageW(pr->hwnd, kUwmPaintAgain, 0, 0);

        CloseHandle(th);
        DestroyTempAllocator();
        return 0;
    }
};

static LRESULT OnPaint(HWND hwnd) {
    Rect rect = ClientRect(hwnd);
    logf("OnPaint: clientRect=%dx%d\n", rect.dx, rect.dy);
    DoubleBuffer buffer(hwnd, rect);
    HDC hdc = buffer.GetDC();
    HBRUSH brushBg = CreateSolidBrush(kColWindowBg);
    HBRUSH brushWhite = GetStockBrush(WHITE_BRUSH);
    RECT rcClient = ToRECT(rect);
    FillRect(hdc, &rcClient, brushBg);

    PdfPreview* preview = (PdfPreview*)GetWindowLongPtr(hwnd, GWLP_USERDATA);
    if (preview && preview->renderer) {
        int pageNo = GetScrollPos(hwnd, SB_VERT);
        RectF page = preview->renderer->GetPageRect(pageNo);
        logf("OnPaint: pageNo=%d, pageRect=%.1fx%.1f\n", pageNo, page.dx, page.dy);
        if (!page.IsEmpty()) {
            rect.Inflate(-kPreviewMargin, -kPreviewMargin);
            float zoom = (float)std::min(rect.dx / page.dx, rect.dy / page.dy) - 0.001f;
            Rect onScreen = RectF((float)rect.x, (float)rect.y, (float)page.dx * zoom, (float)page.dy * zoom).Round();
            onScreen.Offset((rect.dx - onScreen.dx) / 2, (rect.dy - onScreen.dy) / 2);
            logf("OnPaint: zoom=%.3f, onScreen=%d,%d,%dx%d\n", zoom, onScreen.x, onScreen.y, onScreen.dx, onScreen.dy);

            RECT rcPage = ToRECT(onScreen);
            FillRect(hdc, &rcPage, brushWhite);
            preview->renderer->Render(hdc, onScreen, pageNo, zoom);
        } else {
            logf("OnPaint: page rect is empty\n");
        }
    } else {
        logf("OnPaint: no preview or no renderer\n");
    }

    DeleteObject(brushBg);
    DeleteObject(brushWhite);

    PAINTSTRUCT ps;
    buffer.Flush(BeginPaint(hwnd, &ps));
    EndPaint(hwnd, &ps);
    return 0;
}

static LRESULT OnVScroll(HWND hwnd, WPARAM wp) {
    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask = SIF_ALL;
    GetScrollInfo(hwnd, SB_VERT, &si);

    switch (LOWORD(wp)) {
        case SB_TOP:
            si.nPos = si.nMin;
            break;
        case SB_BOTTOM:
            si.nPos = si.nMax;
            break;
        case SB_LINEUP:
            si.nPos--;
            break;
        case SB_LINEDOWN:
            si.nPos++;
            break;
        case SB_PAGEUP:
            si.nPos--;
            break;
        case SB_PAGEDOWN:
            si.nPos++;
            break;
        case SB_THUMBTRACK:
            si.nPos = si.nTrackPos;
            break;
    }
    si.fMask = SIF_POS;
    SetScrollInfo(hwnd, SB_VERT, &si, TRUE);

    InvalidateRect(hwnd, nullptr, TRUE);
    UpdateWindow(hwnd);
    return 0;
}

static LRESULT OnKeydown(HWND hwnd, WPARAM key) {
    switch (key) {
        case VK_DOWN:
        case VK_RIGHT:
        case VK_NEXT:
            return OnVScroll(hwnd, SB_PAGEDOWN);
        case VK_UP:
        case VK_LEFT:
        case VK_PRIOR:
            return OnVScroll(hwnd, SB_PAGEUP);
        case VK_HOME:
            return OnVScroll(hwnd, SB_TOP);
        case VK_END:
            return OnVScroll(hwnd, SB_BOTTOM);
        default:
            return 0;
    }
}

static LRESULT OnDestroy(HWND hwnd) {
    PdfPreview* preview = (PdfPreview*)GetWindowLongPtr(hwnd, GWLP_USERDATA);
    if (preview) {
        delete preview->renderer;
        preview->renderer = nullptr;
    }
    return 0;
}

static LRESULT CALLBACK PreviewWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
        case WM_PAINT:
            return OnPaint(hwnd);
        case WM_VSCROLL:
            return OnVScroll(hwnd, wp);
        case WM_KEYDOWN:
            return OnKeydown(hwnd, wp);
        case WM_LBUTTONDOWN:
            HwndSetFocus(hwnd);
            return 0;
        case WM_MOUSEWHEEL: {
            auto delta = GET_WHEEL_DELTA_WPARAM(wp);
            wp = delta > 0 ? SB_LINEUP : SB_LINEDOWN;
            return OnVScroll(hwnd, wp);
        }
        case WM_DESTROY:
            return OnDestroy(hwnd);
        case kUwmPaintAgain:
            InvalidateRect(hwnd, nullptr, TRUE);
            UpdateWindow(hwnd);
            return 0;
        default:
            return DefWindowProc(hwnd, msg, wp, lp);
    }
}

IFACEMETHODIMP PdfPreview::DoPreview() {
    log("PdfPreview::DoPreview()\n");

    WNDCLASSEX wcex{};
    wcex.cbSize = sizeof(wcex);
    wcex.lpfnWndProc = PreviewWndProc;
    wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wcex.lpszClassName = L"SumatraPDF_PreviewPane";
    wcex.style = CS_HREDRAW | CS_VREDRAW;
    RegisterClassEx(&wcex);

    logf("PdfPreview::DoPreview: parent hwnd=%p, rect=%d,%d,%d,%d\n", m_hwndParent, m_rcParent.x, m_rcParent.y,
         m_rcParent.dx, m_rcParent.dy);

    m_hwnd = CreateWindow(wcex.lpszClassName, nullptr, WS_CHILD | WS_VSCROLL | WS_VISIBLE, m_rcParent.x, m_rcParent.y,
                          m_rcParent.dx, m_rcParent.dy, m_hwndParent, nullptr, nullptr, nullptr);
    if (!m_hwnd) {
        logf("PdfPreview::DoPreview: CreateWindow failed, error=%d\n", (int)GetLastError());
        return HRESULT_FROM_WIN32(GetLastError());
    }
    logf("PdfPreview::DoPreview: created window hwnd=%p\n", m_hwnd);

    this->renderer = nullptr;
    SetWindowLongPtr(m_hwnd, GWLP_USERDATA, (LONG_PTR)this);

    int pageCount = 1;
    if (InitPreviewSession() && pipeSession) {
        pageCount = pipeSession->pageCount;
        this->renderer = new PageRenderer(pipeSession, m_hwnd);
        logf("PdfPreview::DoPreview: session ok, pageCount=%d\n", pageCount);
    } else {
        logf("PdfPreview::DoPreview: InitPreviewSession failed\n");
    }

    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask = SIF_ALL;
    si.nPos = 1;
    si.nMin = 1;
    si.nMax = pageCount;
    si.nPage = si.nMax > 1 ? 1 : 2;
    SetScrollInfo(m_hwnd, SB_VERT, &si, TRUE);

    ShowWindow(m_hwnd, SW_SHOW);
    logf("PdfPreview::DoPreview: done, returning S_OK\n");
    return S_OK;
}

// DLL exports

long g_lRefCount = 0;

class PreviewClassFactory : public IClassFactory {
  public:
    explicit PreviewClassFactory(REFCLSID rclsid) : m_lRef(1), m_clsid(rclsid) { InterlockedIncrement(&g_lRefCount); }

    ~PreviewClassFactory() { InterlockedDecrement(&g_lRefCount); }

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) {
        log("PdfPreview: QueryInterface()\n");
        static const QITAB qit[] = {QITABENT(PreviewClassFactory, IClassFactory), {nullptr}};
        return QISearch(this, qit, riid, ppv);
    }

    IFACEMETHODIMP_(ULONG) AddRef() { return InterlockedIncrement(&m_lRef); }

    IFACEMETHODIMP_(ULONG) Release() {
        long cRef = InterlockedDecrement(&m_lRef);
        if (cRef == 0) {
            delete this;
        }
        return cRef;
    }

    bool IsClsid(const char* s) {
        CLSID clsid;
        return SUCCEEDED(CLSIDFromString(s, &clsid)) && IsEqualCLSID(m_clsid, clsid);
    }

    // IClassFactory
    IFACEMETHODIMP CreateInstance(IUnknown* punkOuter, REFIID riid, void** ppv) {
        log("PdfPreview: CreateInstance()\n");

        *ppv = nullptr;
        if (punkOuter) {
            return CLASS_E_NOAGGREGATION;
        }

        PreviewFileType fileType;
        if (IsClsid(kPdfPreview2Clsid)) {
            fileType = PreviewFileType::PDF;
        } else if (IsClsid(kXpsPreview2Clsid)) {
            fileType = PreviewFileType::XPS;
        } else if (IsClsid(kDjVuPreview2Clsid)) {
            fileType = PreviewFileType::DjVu;
        } else if (IsClsid(kEpubPreview2Clsid)) {
            fileType = PreviewFileType::EPUB;
        } else if (IsClsid(kFb2Preview2Clsid)) {
            fileType = PreviewFileType::FB2;
        } else if (IsClsid(kMobiPreview2Clsid)) {
            fileType = PreviewFileType::MOBI;
        } else if (IsClsid(kCbxPreview2Clsid)) {
            fileType = PreviewFileType::CBX;
        } else if (IsClsid(kTgaPreview2Clsid)) {
            fileType = PreviewFileType::TGA;
        } else {
            return E_NOINTERFACE;
        }

        ScopedComPtr<IInitializeWithStream> pObject;
        pObject = new PdfPreview(&g_lRefCount, fileType);

        if (!pObject) {
            return E_OUTOFMEMORY;
        }

        return pObject->QueryInterface(riid, ppv);
    }

    IFACEMETHODIMP LockServer(BOOL bLock) {
        if (bLock) {
            InterlockedIncrement(&g_lRefCount);
        } else {
            InterlockedDecrement(&g_lRefCount);
        }
        return S_OK;
    }

  private:
    long m_lRef;
    CLSID m_clsid;
};

static const char* GetReason(DWORD dwReason) {
    switch (dwReason) {
        case DLL_PROCESS_ATTACH:
            return "DLL_PROCESS_ATTACH";
        case DLL_THREAD_ATTACH:
            return "DLL_THREAD_ATTACH";
        case DLL_THREAD_DETACH:
            return "DLL_THREAD_DETACH";
        case DLL_PROCESS_DETACH:
            return "DLL_PROCESS_DETACH";
    }
    return "Unknown reason";
}

STDAPI_(BOOL) DllMain(HINSTANCE hInstance, DWORD dwReason, void*) {
    if (dwReason == DLL_PROCESS_ATTACH) {
        ReportIf(hInstance != GetInstance());
    }
    gLogAppName = "PdfPreview";
    logf("PdfPreview: DllMain %s\n", GetReason(dwReason));
    return TRUE;
}

STDAPI DllCanUnloadNow(VOID) {
    return g_lRefCount == 0 ? S_OK : S_FALSE;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv) {
    *ppv = nullptr;
    ScopedComPtr<PreviewClassFactory> pClassFactory(new PreviewClassFactory(rclsid));
    if (!pClassFactory) {
        return E_OUTOFMEMORY;
    }
    log("PdfPreview: DllGetClassObject\n");
    return pClassFactory->QueryInterface(riid, ppv);
}

STDAPI DllRegisterServer() {
    TempStr dllPath = GetSelfExePathTemp();
    if (!dllPath) {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    logf("DllRegisterServer: dllPath=%s\n", dllPath);

    // for compat with SumatraPDF 3.3 and lower
    // in 3.4 we call this code from the installer
    // pre-3.4 we would write to both HKLM (if permissions) and HKCU
    // in 3.4+ this will only install for current user (HKCU)
    bool ok = InstallPreviewDll(dllPath, false);
    if (!ok) {
        log("DllRegisterServer failed!\n");
        return E_FAIL;
    }
    return S_OK;
}

STDAPI DllUnregisterServer() {
    log("DllUnregisterServer\n");

    bool ok = UninstallPreviewDll();
    if (!ok) {
        log("DllUnregisterServer failed!\n");
        return E_FAIL;
    }
    return S_OK;
}

// TODO: maybe remove, is anyone using this functionality?
STDAPI DllInstall(BOOL bInstall, const WCHAR* pszCmdLine) {
    char* s = ToUtf8Temp(pszCmdLine);
    DisablePreviewInstallExts(s);
    if (!bInstall) {
        return DllUnregisterServer();
    }
    return DllRegisterServer();
}

#ifdef _WIN64
#pragma comment(linker, "/EXPORT:DllCanUnloadNow=DllCanUnloadNow,PRIVATE")
#pragma comment(linker, "/EXPORT:DllGetClassObject=DllGetClassObject,PRIVATE")
#pragma comment(linker, "/EXPORT:DllRegisterServer=DllRegisterServer,PRIVATE")
#pragma comment(linker, "/EXPORT:DllUnregisterServer=DllUnregisterServer,PRIVATE")
#pragma comment(linker, "/EXPORT:DllInstall=DllInstall,PRIVATE")
#else
#pragma comment(linker, "/EXPORT:DllCanUnloadNow=_DllCanUnloadNow@0,PRIVATE")
#pragma comment(linker, "/EXPORT:DllGetClassObject=_DllGetClassObject@12,PRIVATE")
#pragma comment(linker, "/EXPORT:DllRegisterServer=_DllRegisterServer@0,PRIVATE")
#pragma comment(linker, "/EXPORT:DllUnregisterServer=_DllUnregisterServer@0,PRIVATE")
#pragma comment(linker, "/EXPORT:DllInstall=_DllInstall@8,PRIVATE")
#endif
