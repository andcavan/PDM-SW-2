#!/usr/bin/env python
"""
Subprocess helper: genera una thumbnail da un file SolidWorks.

Strategia: usa IThumbnailProvider via COM vtable, invocando direttamente
il thumbnail handler di eDrawings ({72670837-AA64-4C1D-AB58-A9D9D31A1216})
già registrato nel sistema. Non richiede apertura di finestre.

Uso: python thumb_worker.py <src_file> <dest_file>
Uscita: 0 = thumbnail salvata, 1 = fallimento
"""
import sys
import time
from pathlib import Path


def main():
    if len(sys.argv) != 3:
        print("Usage: thumb_worker.py <src_file> <dest_file>", file=sys.stderr)
        sys.exit(2)

    src  = Path(sys.argv[1])
    dest = Path(sys.argv[2])

    if not src.exists():
        print(f"File non trovato: {src}", file=sys.stderr)
        sys.exit(1)

    # Dimensione di estrazione: lato lungo 400 px (i DRW sono orizzontali/verticali)
    # Qt ridimensionerà in proporzione al contenitore 1:1.5.
    if _try_thumbnail_provider(src, dest, size=400):
        print(f"OK (IThumbnailProvider): {dest.name}")
        sys.exit(0)

    print(f"FAIL: impossibile generare thumbnail per {src.name}", file=sys.stderr)
    sys.exit(1)


# =============================================================================
#  Approccio 1 – IThumbnailProvider via COM vtable
# =============================================================================

def _try_thumbnail_provider(src: Path, dest: Path, size: int = 256) -> bool:
    """
    Invoca il thumbnail handler di eDrawings tramite IThumbnailProvider.
    CLSID: {72670837-AA64-4C1D-AB58-A9D9D31A1216}
    Interfacce: IInitializeWithWindow (opzionale) + IInitializeWithFile + IThumbnailProvider
    """
    import ctypes
    import ctypes.wintypes as wt
    import uuid
    import pythoncom
    from PIL import Image

    ole32    = ctypes.windll.ole32
    gdi32    = ctypes.windll.gdi32
    user32   = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32

    # Se il processo non ha una console (es. avviato da pythonw.exe via subprocess),
    # il renderer eDrawings non riesce a creare il contesto GUI necessario.
    # Allociamo una console nascosta per fornire il contesto richiesto.
    _console_allocated = False
    if not kernel32.GetConsoleWindow():
        if kernel32.AllocConsole():
            _console_allocated = True
            # Nascondi immediatamente la console
            hwnd_con = kernel32.GetConsoleWindow()
            if hwnd_con:
                user32.ShowWindow(hwnd_con, 0)  # SW_HIDE

    class GUID(ctypes.Structure):
        _fields_ = [
            ('Data1', wt.DWORD),
            ('Data2', wt.WORD),
            ('Data3', wt.WORD),
            ('Data4', ctypes.c_byte * 8),
        ]

    def make_guid(s: str) -> GUID:
        b = uuid.UUID(s).bytes_le
        g = GUID()
        ctypes.memmove(ctypes.byref(g), b, 16)
        return g

    # Costanti COM
    CLSID_THUMB     = make_guid('{72670837-AA64-4C1D-AB58-A9D9D31A1216}')
    IID_INIT_FILE   = make_guid('{B7D14566-0509-4CCE-A71F-0A554233BD9B}')
    IID_THUMB_PROV  = make_guid('{E357FCCD-A995-4576-B01F-234630154E96}')
    IID_INIT_WINDOW = make_guid('{3D73A659-E5D0-4D42-AFC0-5121BA425C8D}')
    CLSCTX_INPROC   = 1
    STGM_READ       = 0

    HRESULT = ctypes.HRESULT

    def vtcall(obj_ptr, idx, restype, *argtypes):
        """Chiama il metodo idx-esimo della vtable dell'oggetto COM."""
        vtable = ctypes.cast(obj_ptr, ctypes.POINTER(ctypes.c_void_p))[0]
        fn_ptr  = ctypes.cast(vtable, ctypes.POINTER(ctypes.c_void_p))[idx]
        proto   = ctypes.WINFUNCTYPE(restype, ctypes.c_void_p, *argtypes)
        return proto(fn_ptr)

    def com_release(obj_ptr):
        try:
            if obj_ptr and obj_ptr.value:
                vtcall(obj_ptr, 2, ctypes.c_ulong)(obj_ptr)
        except Exception:
            pass

    # Crea una finestra nascosta da fornire al renderer COM come contesto GUI
    hwnd = user32.CreateWindowExW(0, "STATIC", "thumb", 0, 0, 0, 1, 1,
                                  None, None, None, None)

    try:
        # MTA: COM crea automaticamente un thread host STA con message pump
        pythoncom.CoInitializeEx(0)   # 0 = COINIT_MULTITHREADED

        # --- CoCreateInstance → IInitializeWithFile ---
        obj = ctypes.c_void_p()
        hr = ole32.CoCreateInstance(
            ctypes.byref(CLSID_THUMB),
            None,
            CLSCTX_INPROC,
            ctypes.byref(IID_INIT_FILE),
            ctypes.byref(obj),
        )
        if hr != 0 or not obj.value:
            print(f"  CoCreateInstance IInitWithFile hr={hr:#010x}", file=sys.stderr)
            # Prova direttamente con IThumbnailProvider
            hr = ole32.CoCreateInstance(
                ctypes.byref(CLSID_THUMB),
                None,
                CLSCTX_INPROC,
                ctypes.byref(IID_THUMB_PROV),
                ctypes.byref(obj),
            )
            if hr != 0 or not obj.value:
                print(f"  CoCreateInstance IThumbnailProvider hr={hr:#010x}", file=sys.stderr)
                pythoncom.CoUninitialize()
                return False
            thumb_obj = obj
        else:
            # --- Opzionale: IInitializeWithWindow per dare contesto GUI al renderer ---
            if hwnd:
                try:
                    iw_obj = ctypes.c_void_p()
                    # Usa c_long per evitare auto-raise su E_NOINTERFACE
                    qi_fn2 = vtcall(obj, 0, ctypes.c_long,
                                    ctypes.POINTER(GUID), ctypes.POINTER(ctypes.c_void_p))
                    hr_iw = qi_fn2(obj, ctypes.byref(IID_INIT_WINDOW), ctypes.byref(iw_obj))
                    if hr_iw == 0 and iw_obj.value:
                        iw_fn = vtcall(iw_obj, 3, HRESULT, wt.HWND)
                        iw_fn(iw_obj, hwnd)
                        com_release(iw_obj)
                except Exception:
                    pass

            # --- IInitializeWithFile::Initialize(path, STGM_READ) ---
            init_fn = vtcall(obj, 3, HRESULT, ctypes.c_wchar_p, wt.DWORD)
            hr = init_fn(obj, str(src), STGM_READ)
            if hr != 0:
                print(f"  IInitializeWithFile::Initialize hr={hr:#010x}", file=sys.stderr)
                com_release(obj)
                pythoncom.CoUninitialize()
                return False

            # --- QI per IThumbnailProvider ---
            thumb_obj = ctypes.c_void_p()
            qi_fn = vtcall(obj, 0, HRESULT,
                           ctypes.POINTER(GUID), ctypes.POINTER(ctypes.c_void_p))
            hr = qi_fn(obj, ctypes.byref(IID_THUMB_PROV), ctypes.byref(thumb_obj))
            com_release(obj)
            if hr != 0 or not thumb_obj.value:
                print(f"  QI IThumbnailProvider hr={hr:#010x}", file=sys.stderr)
                pythoncom.CoUninitialize()
                return False

        try:
            # --- IThumbnailProvider::GetThumbnail(cx, phbmp, pdwAlpha) ---
            hbmp  = wt.HBITMAP()
            alpha = wt.DWORD()
            gt_fn = vtcall(thumb_obj, 3, HRESULT,
                           wt.UINT,
                           ctypes.POINTER(wt.HBITMAP),
                           ctypes.POINTER(wt.DWORD))
            hr = gt_fn(thumb_obj, size, ctypes.byref(hbmp), ctypes.byref(alpha))
            if hr != 0 or not hbmp.value:
                print(f"  GetThumbnail hr={hr:#010x}", file=sys.stderr)
                return False

            # --- HBITMAP → PIL Image ---
            img = _hbitmap_to_pil(hbmp.value, gdi32)
            if img is None:
                return False

            dest.parent.mkdir(parents=True, exist_ok=True)
            img.thumbnail((size, size), Image.LANCZOS)
            img.save(str(dest), 'PNG')
            return True

        finally:
            com_release(thumb_obj)
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    except Exception as e:
        print(f"  IThumbnailProvider: {e}", file=sys.stderr)
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        return False
    finally:
        if hwnd:
            try:
                user32.DestroyWindow(hwnd)
            except Exception:
                pass
        if _console_allocated:
            try:
                kernel32.FreeConsole()
            except Exception:
                pass


def _hbitmap_to_pil(hbmp, gdi32):
    """Converte un HBITMAP in un'immagine PIL usando win32ui."""
    try:
        import win32ui
        import win32con
        from PIL import Image

        bmp = win32ui.CreateBitmapFromHandle(hbmp)
        info = bmp.GetInfo()
        W, H = info['bmWidth'], info['bmHeight']
        if W <= 0 or H <= 0:
            return None

        bits = bmp.GetBitmapBits(True)
        # BGRA → RGB
        img = Image.frombuffer('RGBA', (W, H), bits, 'raw', 'BGRA', 0, 1)
        return img.convert('RGB')

    except Exception as e:
        print(f"  HBITMAP→PIL: {e}", file=sys.stderr)
        return None
    finally:
        if hbmp:
            import ctypes
            # Su Windows 64-bit HBITMAP può essere > 2^31: usare c_void_p
            gdi32.DeleteObject(ctypes.c_void_p(hbmp))


# =============================================================================
#  Approccio 2 – eDrawings COM (MTA)
# =============================================================================

def _try_edrawings(src: Path, dest: Path) -> bool:
    """Genera thumbnail via eDrawings COM in modalità MTA."""
    try:
        import pythoncom
        import win32com.client

        pythoncom.CoInitializeEx(0)   # COINIT_MULTITHREADED
        try:
            eApp = win32com.client.Dispatch("EModelView.EModelViewControl")
            eApp.OpenDoc(str(src), False, False, True, "")

            deadline = time.time() + 20
            loaded = False
            while time.time() < deadline:
                time.sleep(0.2)
                try:
                    if eApp.FileName:
                        loaded = True
                        break
                except Exception:
                    pass

            if loaded:
                dest.parent.mkdir(parents=True, exist_ok=True)
                eApp.SaveAs(str(dest))
                for _ in range(50):
                    time.sleep(0.2)
                    if dest.exists():
                        eApp.CloseActiveDoc("")
                        return True
            try:
                eApp.CloseActiveDoc("")
            except Exception:
                pass

        finally:
            pythoncom.CoUninitialize()

    except Exception as e:
        print(f"  eDrawings: {e}", file=sys.stderr)

    return False


if __name__ == '__main__':
    main()
