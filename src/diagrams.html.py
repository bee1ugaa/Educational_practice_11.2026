# Генерация HTML-страниц архитектурных схем (рендер в PNG выполняет render_diagrams.js)
import sys, pathlib
out = pathlib.Path(sys.argv[1])
out.mkdir(parents=True, exist_ok=True)

CSS = """
*{box-sizing:border-box;margin:0;padding:0}
body{width:1600px;background:#fff;font-family:'Liberation Serif','Times New Roman',serif;color:#000}
.frame{border:3px solid #000;padding:26px 30px 30px}
.title{font-size:40px;font-weight:bold;text-align:center;margin-bottom:18px}
.zone{border:3px solid #000;border-radius:6px;padding:14px 16px 16px;position:relative}
.zone.user{background:#f4f4f4}
.zone.kernel{background:#dcdcdc}
.zlabel{font-size:30px;font-weight:bold;margin-bottom:10px}
.row{display:flex;gap:14px;margin-top:12px}
.row:first-of-type{margin-top:0}
.b{flex:1;background:#fff;border:2px solid #000;border-radius:4px;padding:10px 12px;font-size:28px;line-height:1.2;text-align:center;display:flex;flex-direction:column;justify-content:center}
.b b{font-size:29px}
.b small{font-size:24px;display:block;margin-top:4px}
.b.dark{background:#bdbdbd}
.b.hatch{background:repeating-linear-gradient(45deg,#fff,#fff 10px,#ececec 10px,#ececec 20px)}
.boundary{margin:14px 0;border-top:4px dashed #000;position:relative;height:0}
.blabel{position:relative;top:-26px;margin:0 auto;width:max-content;background:#fff;padding:4px 18px;font-size:27px;border:2px solid #000;border-radius:4px}
.arrows{display:flex;justify-content:space-around;font-size:28px;line-height:1;height:40px;align-items:center}
.hw{margin-top:12px;border:3px solid #000;background:#8c8c8c;color:#fff;font-size:30px;font-weight:bold;text-align:center;padding:12px;border-radius:6px}
.legend{margin-top:14px;font-size:24px;display:flex;gap:36px;justify-content:center;align-items:center}
.sw{display:inline-block;width:34px;height:22px;border:2px solid #000;vertical-align:middle;margin-right:8px}
"""

def page(name, title, body, legend):
    html = f"""<!doctype html><html><head><meta charset="utf-8"><style>{CSS}</style></head>
<body><div class="frame" id="cap">{body}
<div class="legend">{legend}</div></div></body></html>"""
    (out / f"{name}.html").write_text(html, encoding="utf-8")

def B(t, s="", cls="", flex=1):
    small = f"<small>{s}</small>" if s else ""
    return f'<div class="b {cls}" style="flex:{flex}"><b>{t}</b>{small}</div>'

def row(*bs):
    return '<div class="row">' + "".join(bs) + "</div>"

ARR = '<div class="arrows"><span>⇅</span><span>⇅</span><span>⇅</span></div>'
LEG_STD = ('<span><span class="sw" style="background:#f4f4f4"></span>пространство пользователя (кольцо 3)</span>'
           '<span><span class="sw" style="background:#dcdcdc"></span>пространство ядра (кольцо 0)</span>'
           )

# ---------- Windows NT ----------
win = f"""
<div class="zone user"><div class="zlabel">Пространство пользователя (user mode)</div>
{row(B("Приложения", "Win32, .NET, UWP"),
     B("Системные процессы", "smss, csrss, wininit, winlogon, lsass"),
     B("Службы", "services.exe (SCM), svchost.exe"),
     B("Подсистемы окружения", "Win32 (csrss), WSL"))}
{row(B("DLL подсистем", "kernel32, kernelbase, user32, gdi32, advapi32", flex=3),
     B("UMDF-драйверы", "процесс WUDFHost.exe", cls="hatch"))}
{row(B("ntdll.dll", "заглушки системных вызовов Native API"))}
</div>
<div class="boundary"><div class="blabel">граница режимов: системный вызов (syscall) → диспетчер системных служб</div></div>
<div class="zone kernel"><div class="zlabel">Пространство ядра (kernel mode) – общее адресное пространство</div>
{row(B("Исполнительная система (Executive), ntoskrnl.exe",
       "диспетчеры объектов, ввода-вывода, памяти, процессов и потоков, кэша, PnP и электропитания; "
       "монитор безопасности (SRM); диспетчер конфигурации (реестр); ALPC"))}
{row(B("Win32k.sys", "оконный менеджер, GDI"),
     B("Драйверы устройств и файловых систем (*.sys)", "WDM/KMDF, NTFS, ReFS, сетевой стек, фильтры", cls="hatch", flex=2.4))}
{row(B("Ядро (Kernel)", "планирование потоков, прерывания, синхронизация"),
     B("HAL (hal.dll)", "уровень аппаратных абстракций", cls="dark"))}
</div>
<div class="hw">Аппаратное обеспечение (процессор, память, устройства)</div>"""
page("windows_nt", "Архитектура Windows NT (Windows 11, NT 10.0) – гибридное ядро", win,
     LEG_STD + '<span><span class="sw" style="background:repeating-linear-gradient(45deg,#fff,#fff 5px,#ccc 5px,#ccc 10px)"></span>драйверы</span>')

# ---------- Linux ----------
lin = f"""
<div class="zone user"><div class="zlabel">Пространство пользователя (user space)</div>
{row(B("Приложения", "браузер, офис, серверное ПО"),
     B("Системные службы", "systemd, journald, udevd, sshd, NetworkManager"),
     B("Оболочка и утилиты", "bash, coreutils"),
     B("Драйверы и ФС в user space", "FUSE, libusb, UIO/VFIO", cls="hatch"))}
{row(B("Стандартная библиотека C (glibc) и системные библиотеки"))}
</div>
<div class="boundary"><div class="blabel">граница: интерфейс системных вызовов (syscall), /proc, /sys, ioctl, netlink</div></div>
<div class="zone kernel"><div class="zlabel">Пространство ядра (kernel space) – монолитное модульное ядро, единое адресное пространство</div>
{row(B("Управление процессами", "планировщик EEVDF, сигналы, cgroups, namespaces"),
     B("Управление памятью", "страничная память, buddy, SLUB, page cache"),
     B("VFS", "виртуальная файловая система"),
     B("Сетевой стек", "TCP/IP, netfilter"))}
{row(B("Файловые системы", "ext4, Btrfs, XFS, NFS"),
     B("IPC", "каналы, сокеты, futex, разделяемая память"),
     B("Безопасность", "LSM: AppArmor, SELinux; capabilities; seccomp"),
     B("eBPF", "проверяемые программы в ядре"))}
{row(B("Драйверы устройств и загружаемые модули ядра (*.ko)", "блочные, сетевые, символьные, GPU (DRM)", cls="hatch"))}
{row(B("Архитектурно-зависимый код (arch/x86, arm64)", "обработка прерываний, переключение контекста, MMU", cls="dark"))}
</div>
<div class="hw">Аппаратное обеспечение (процессор, память, устройства)</div>"""
page("linux", "Архитектура ОС на базе Linux (Ubuntu 24.04 LTS, ядро 6.8) – монолитное модульное ядро", lin,
     LEG_STD + '<span><span class="sw" style="background:repeating-linear-gradient(45deg,#fff,#fff 5px,#ccc 5px,#ccc 10px)"></span>драйверы</span>')

# ---------- MINIX 3 ----------
mnx = f"""
<div class="zone user"><div class="zlabel">Пространство пользователя – каждый компонент является отдельным процессом</div>
{row(B("Уровень 4: пользовательские процессы", "sh, утилиты и библиотеки NetBSD, init, приложения"))}
<div class="arrows"><span>⇅ сообщения</span><span>⇅ сообщения</span></div>
{row(B("PM", "процессы, сигналы"), B("VM", "виртуальная память"), B("SCHED", "политика планирования"),
     B("VFS", "вирт. ФС"), B("MFS / ext2", "файловые системы"), B("INET (LWIP)", "сеть"), B("DS", "хранилище данных"),
     B("RS", "сервер реинкарнации", cls="dark"))}
<div style="font-size:26px;text-align:center;margin:6px 0 0">Уровень 3: системные серверы</div>
<div class="arrows"><span>⇅ сообщения</span><span>⇅ сообщения</span></div>
{row(B("Драйвер диска", "AHCI, ATA", cls="hatch"), B("Драйвер TTY", "консоль", cls="hatch"),
     B("Драйвер Ethernet", "e1000, virtio", cls="hatch"), B("Аудио, USB, PCI", "", cls="hatch"),
     B("…", "прочие драйверы", cls="hatch"))}
<div style="font-size:26px;text-align:center;margin:6px 0 0">Уровень 2: драйверы устройств (без привилегий кольца 0, доступ к памяти через гранты)</div>
</div>
<div class="boundary"><div class="blabel">граница: IPC ядра (send, receive, sendrec, notify, senda) и вызовы ядра sys_* по маске привилегий</div></div>
<div class="zone kernel"><div class="zlabel">Пространство ядра – микроядро (несколько тысяч строк кода)</div>
{row(B("Микроядро", "прерывания, переключение процессов, передача сообщений, MMU", flex=2),
     B("Задача SYSTEM", "обслуживание вызовов ядра"),
     B("Задача CLOCK", "таймеры, учёт времени"))}
</div>
<div class="hw">Аппаратное обеспечение (процессор, память, устройства)</div>"""
page("minix3", "Архитектура MINIX 3 (версия 3.4.0) – микроядро с многосерверной организацией", mnx,
     LEG_STD + '<span>⇅ – обмен сообщениями через ядро</span>' + '<span><span class="sw" style="background:repeating-linear-gradient(45deg,#fff,#fff 5px,#ccc 5px,#ccc 10px)"></span>драйверы</span>')
print("ok")
