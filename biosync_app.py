"""
BioSync - Attendance Sync
Frappe-style UI | Calendar datetime picker | Windows Task Scheduler
"""
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkcalendar import Calendar
import json, os, threading, subprocess, sys
from datetime import datetime

CONFIG_FILE = os.path.join(os.path.expanduser("~"), "BioSync", "biosync_config.json")
LOG_FILE    = os.path.join(os.path.expanduser("~"), "BioSync", "biosync.log")
TASK_NAME   = "BioSyncScheduler"
os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)

BG      = "#f4f5f6"
WHITE   = "#ffffff"
BORDER  = "#d1d8dd"
PRIMARY = "#2490ef"
DANGER  = "#e03636"
SUCCESS = "#28a745"
TEXT    = "#1c2126"
MUTED   = "#6c7680"
LABEL   = "#8d99a6"
NAV_BG  = "#2c3e50"
FF      = "Segoe UI"
def F(s=10, w="normal"): return (FF, s, w)

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {"folder_path":"","file_name":"eTimeTrackLite1.mdb",
            "schedule_type":"hour","schedule_value":"1",
            "prev_synced_date":"","url":"","username":"","password":""}

def save_config(cfg):
    with open(CONFIG_FILE,"w", encoding="utf-8") as f: json.dump(cfg,f,indent=4)

def write_log(msg, box=None):
    ts   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}\n"
    with open(LOG_FILE,"a", encoding="utf-8") as f: f.write(line)
    if box:
        try:
            box.config(state="normal")
            box.insert("end",line)
            box.see("end")
            box.config(state="disabled")
        except: pass

def fetch_punches(mdb_path, from_dt=None, to_dt=None, log=None):
    try:
        import pyodbc
    except ImportError:
        write_log("pyodbc not installed. Run: pip install pyodbc", log); return []
    if not os.path.exists(mdb_path):
        write_log(f"File not found: {mdb_path}", log); return []
    try:
        conn   = pyodbc.connect(r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq="+mdb_path+";")
        cursor = conn.cursor()
        device_map = {}
        try:
            cursor.execute("SELECT DeviceId, DeviceFName FROM Devices")
            for r in cursor.fetchall():
                device_map[str(int(r[0])) if r[0] else ""] = str(r[1]).strip()
        except: pass
        tables = [r.table_name for r in cursor.tables(tableType="TABLE")
                  if r.table_name.lower().startswith("devicelogs")]
        write_log(f"Tables: {tables}", log)
        punches = []
        for t in tables:
            try:
                if from_dt and to_dt:
                    cursor.execute(f"SELECT UserId,LogDate,DeviceId FROM [{t}] WHERE LogDate>? AND LogDate<=?",from_dt,to_dt)
                elif from_dt:
                    cursor.execute(f"SELECT UserId,LogDate,DeviceId FROM [{t}] WHERE LogDate>?",from_dt)
                else:
                    cursor.execute(f"SELECT UserId,LogDate,DeviceId FROM [{t}]")
                rows = cursor.fetchall()
                for r in rows:
                    did = str(int(r[2])) if r[2] else ""
                    punches.append({"employee_id":str(r[0]).strip(),
                                    "time":r[1].strftime("%Y-%m-%d %H:%M:%S") if r[1] else "",
                                    "device_name":device_map.get(did,f"device_{did}")})
                write_log(f"  {t}: {len(rows)} records",log)
            except Exception as e:
                write_log(f"  {t} error: {e}",log)
        conn.close()
        return punches
    except Exception as e:
        write_log(f"DB error: {e}",log); return []

def push_data(url, user, pwd, punches, log=None):
    try:
        import requests
        write_log(f"Sending {len(punches)} records to server...", log)
        r = requests.post(url, json={"data": punches},
                          auth=(user, pwd), timeout=60)
        write_log(f"HTTP {r.status_code}: {r.text[:200]}", log)
        return r.status_code in (200, 201)
    except Exception as e:
        write_log(f"Push error: {e}", log)
        return False

def get_exe_path():
    return sys.executable if getattr(sys,'frozen',False) else os.path.abspath(__file__)

def register_task(stype, sval, log=None):
    exe    = get_exe_path()
    is_exe = getattr(sys,'frozen',False)
    run_cmd = f'"{exe}" --sync' if is_exe else f'"{sys.executable}" "{exe}" --sync'
    subprocess.run(["schtasks","/Delete","/TN",TASK_NAME,"/F"],capture_output=True,shell=True)
    base = ["schtasks","/Create","/TN",TASK_NAME,"/TR",run_cmd,"/RL","HIGHEST","/F"]
    cmds = {
        "minute": base+["/SC","MINUTE","/MO",str(sval)],
        "hour"  : base+["/SC","HOURLY","/MO",str(sval)],
        "day"   : base+["/SC","DAILY", "/MO",str(sval),"/ST","00:00"],
        "month" : base+["/SC","MONTHLY","/MO",str(sval),"/ST","00:00"],
    }
    cmd = cmds.get(stype)
    if not cmd: return False
    result = subprocess.run(cmd,capture_output=True,text=True,shell=True)
    ok = result.returncode == 0
    write_log(f"{'✅' if ok else '❌'} Task Scheduler: every {sval} {stype}(s) — {result.stderr if not ok else 'OK'}",log)
    return ok

def remove_task(log=None):
    r = subprocess.run(["schtasks","/Delete","/TN",TASK_NAME,"/F"],capture_output=True,text=True,shell=True)
    write_log("✅ Task removed" if r.returncode==0 else "No task found",log)

# ── Calendar DateTime Dialog ──────────────────────────────────────────────────

class DateTimeDialog(tk.Toplevel):
    def __init__(self, parent, title="Select Date & Time", initial=None):
        super().__init__(parent)
        self.title(title)
        self.resizable(False,False)
        self.configure(bg=WHITE)
        self.result = None
        self.grab_set()
        self.transient(parent)
        dt = initial or datetime.now()

        tk.Label(self,text="Select Date",bg=WHITE,fg=MUTED,font=F(9,"bold")).pack(padx=16,pady=(16,4),anchor="w")
        self.cal = Calendar(self,selectmode="day",
                            year=dt.year,month=dt.month,day=dt.day,
                            date_pattern="yyyy-mm-dd",
                            background=PRIMARY,foreground=WHITE,
                            headersbackground=NAV_BG,headersforeground=WHITE,
                            selectbackground=PRIMARY,
                            normalbackground=WHITE,
                            weekendbackground="#f0f4f7",
                            othermonthbackground="#e8ecef",
                            font=F(10),borderwidth=0)
        self.cal.pack(padx=16,pady=(0,8))

        tk.Label(self,text="Select Time",bg=WHITE,fg=MUTED,font=F(9,"bold")).pack(padx=16,pady=(8,4),anchor="w")
        tf = tk.Frame(self,bg=WHITE)
        tf.pack(padx=16,pady=(0,16))

        def sp(label,fr,to,val):
            tk.Label(tf,text=label,bg=WHITE,fg=LABEL,font=F(9)).pack(side="left",padx=(0,2))
            s = tk.Spinbox(tf,from_=fr,to=to,width=4,format="%02.0f",
                           font=F(11),bg="#f7f9fb",fg=TEXT,
                           relief="solid",bd=1,buttonbackground=BG)
            s.delete(0,"end"); s.insert(0,f"{val:02d}")
            s.pack(side="left",ipady=6,padx=(0,10))
            return s

        self.h = sp("Hour",0,23,dt.hour)
        self.m = sp("Min", 0,59,dt.minute)
        self.s = sp("Sec", 0,59,dt.second)

        br = tk.Frame(self,bg=WHITE)
        br.pack(fill="x",padx=16,pady=(0,16))
        tk.Button(br,text="✔  Select",bg=PRIMARY,fg=WHITE,
                  relief="flat",font=F(10,"bold"),cursor="hand2",
                  padx=16,pady=8,command=self._pick).pack(side="left",padx=(0,8))
        tk.Button(br,text="Cancel",bg=BG,fg=MUTED,
                  relief="flat",font=F(10),cursor="hand2",
                  padx=16,pady=8,command=self.destroy).pack(side="left")
        self.wait_window()

    def _pick(self):
        try:
            self.result = datetime.strptime(
                f"{self.cal.get_date()} {int(self.h.get()):02d}:{int(self.m.get()):02d}:{int(self.s.get()):02d}",
                "%Y-%m-%d %H:%M:%S")
            self.destroy()
        except Exception as e:
            messagebox.showerror("Error",str(e))

class DateTimeField(tk.Frame):
    def __init__(self, parent, bg=WHITE, initial=None, **kw):
        super().__init__(parent,bg=bg,**kw)
        self._val = initial or datetime.now()
        self.disp = tk.Entry(self,font=F(10),bg="#f7f9fb",fg=TEXT,
                             relief="solid",bd=1,state="readonly",
                             readonlybackground="#f7f9fb",width=22)
        self.disp.pack(side="left",ipady=6,padx=(0,6))
        tk.Button(self,text="📅 Pick",bg=WHITE,fg=PRIMARY,
                  relief="solid",bd=1,font=F(9),cursor="hand2",
                  padx=8,pady=4,command=self._open).pack(side="left")
        self._refresh()

    def _refresh(self):
        self.disp.config(state="normal")
        self.disp.delete(0,"end")
        self.disp.insert(0,self._val.strftime("%Y-%m-%d  %H:%M:%S"))
        self.disp.config(state="readonly")

    def _open(self):
        d = DateTimeDialog(self.winfo_toplevel(),initial=self._val)
        if d.result:
            self._val = d.result
            self._refresh()

    def get_dt(self): return self._val

    def set_dt(self, s):
        try:
            self._val = datetime.strptime(s,"%Y-%m-%d %H:%M:%S")
            self._refresh()
        except: pass

# ── App ───────────────────────────────────────────────────────────────────────

class App:
    def __init__(self, root):
        self.root = root
        self.cfg  = load_config()
        root.title("BioSync — Attendance Sync")
        root.geometry("780x900")
        root.configure(bg=BG)
        root.resizable(True,True)
        root.minsize(700,700)
        self._build()
        self._load_fields()
        write_log("BioSync started",self.log_box)

    def _build(self):
        nav = tk.Frame(self.root,bg=NAV_BG,height=52)
        nav.pack(fill="x"); nav.pack_propagate(False)
        tk.Label(nav,text="  BioSync",bg=NAV_BG,fg=WHITE,font=F(15,"bold")).pack(side="left",padx=14)
        self.status_lbl = tk.Label(nav,text="⬤  Idle",bg=NAV_BG,fg="#7f8c8d",font=F(9))
        self.status_lbl.pack(side="right",padx=20)

        tabbar = tk.Frame(self.root,bg=WHITE,highlightthickness=1,highlightbackground=BORDER)
        tabbar.pack(fill="x")
        content = tk.Frame(self.root,bg=BG)
        content.pack(fill="both",expand=True)

        self.pages={}; self.tab_btns={}
        for key,lbl in [("settings","⚙  Settings"),("logs","📋  Logs")]:
            pg = tk.Frame(content,bg=BG)
            pg.place(relwidth=1,relheight=1)
            self.pages[key]=pg
            b = tk.Button(tabbar,text=lbl,bg=WHITE,fg=MUTED,
                          relief="flat",bd=0,padx=20,pady=10,
                          font=F(10),cursor="hand2",
                          command=lambda k=key:self._show(k))
            b.pack(side="left"); self.tab_btns[key]=b

        self._build_settings(self.pages["settings"])
        self._build_logs(self.pages["logs"])
        self._show("settings")

    def _show(self,key):
        self.pages[key].lift()
        for k,b in self.tab_btns.items():
            b.config(fg=PRIMARY if k==key else MUTED,
                     font=F(10,"bold") if k==key else F(10))

    def _build_settings(self,parent):
        canvas = tk.Canvas(parent,bg=BG,highlightthickness=0)
        vsb = tk.Scrollbar(parent,orient="vertical",command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right",fill="y"); canvas.pack(fill="both",expand=True)
        inner = tk.Frame(canvas,bg=BG)
        cwin  = canvas.create_window((0,0),window=inner,anchor="nw")
        canvas.bind("<Configure>",lambda e:canvas.itemconfig(cwin,width=e.width))
        inner.bind("<Configure>",lambda e:canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind_all("<MouseWheel>",lambda e:canvas.yview_scroll(int(-1*(e.delta/120)),"units"))

        def sec(title,icon):
            hf = tk.Frame(inner,bg=BG)
            hf.pack(fill="x",padx=24,pady=(20,0))
            tk.Label(hf,text=f"{icon}  {title}",bg=BG,fg=MUTED,font=F(9,"bold")).pack(side="left")
            tk.Frame(hf,bg=BORDER,height=1).pack(side="left",fill="x",expand=True,padx=(10,0))
            c = tk.Frame(inner,bg=WHITE,highlightthickness=1,highlightbackground=BORDER)
            c.pack(fill="x",padx=24,pady=(6,0))
            return c

        def row(c,label):
            r = tk.Frame(c,bg=WHITE)
            r.pack(fill="x",padx=16,pady=10)
            tk.Label(r,text=label,bg=WHITE,fg=LABEL,font=F(9),width=22,anchor="w").pack(side="left")
            return r

        def inp(r,var,show=""):
            e = tk.Entry(r,textvariable=var,show=show,font=F(10),
                         bg="#f7f9fb",fg=TEXT,relief="solid",bd=1,
                         highlightthickness=1,highlightcolor=PRIMARY,highlightbackground=BORDER)
            e.pack(side="left",fill="x",expand=True,ipady=6)
            e.bind("<FocusIn>", lambda ev:e.config(highlightbackground=PRIMARY,bg=WHITE))
            e.bind("<FocusOut>",lambda ev:e.config(highlightbackground=BORDER,bg="#f7f9fb"))
            return e

        # DATABASE
        c1 = sec("DATABASE FILE","📁")
        r1 = row(c1,"Folder Path"); self.v_folder=tk.StringVar(); inp(r1,self.v_folder)
        tk.Button(r1,text="Browse",bg=BG,fg=PRIMARY,relief="solid",bd=1,
                  font=F(9),padx=10,pady=4,cursor="hand2",
                  command=lambda:self.v_folder.set(filedialog.askdirectory() or self.v_folder.get())
                  ).pack(side="left",padx=(8,0))
        r2 = row(c1,"File Name"); self.v_file=tk.StringVar(); inp(r2,self.v_file)

        # REMOTE SERVER
        c2 = sec("REMOTE SERVER","🌐")
        r3=row(c2,"Endpoint URL");    self.v_url =tk.StringVar(); inp(r3,self.v_url)
        r4=row(c2,"Username/API Key");self.v_user=tk.StringVar(); inp(r4,self.v_user)
        r5=row(c2,"Password/Secret"); self.v_pass=tk.StringVar(); inp(r5,self.v_pass,show="●")

        # SCHEDULER
        c3 = sec("SCHEDULER","⏱")

        rs = row(c3,"Schedule Type")
        self.v_stype = tk.StringVar(value="hour")
        for val,lbl in [("minute","Every N Minutes"),("hour","Every N Hours"),
                        ("day","Every N Days"),("month","Every N Months")]:
            tk.Radiobutton(rs,text=lbl,variable=self.v_stype,value=val,
                           bg=WHITE,fg=TEXT,font=F(10),
                           activebackground=WHITE,selectcolor=WHITE,
                           command=self._hint).pack(side="left",padx=(0,12))

        rv = row(c3,"N Value")
        self.v_sval = tk.StringVar(value="1")
        inp(rv,self.v_sval)
        self.hint_lbl = tk.Label(rv,text="e.g. 1 = every 1 hour",
                                 bg=WHITE,fg=LABEL,font=F(8))
        self.hint_lbl.pack(side="left",padx=(8,0))

        rp = row(c3,"Previous Synced Date")
        self.v_prev = tk.StringVar()
        pe = tk.Entry(rp,textvariable=self.v_prev,state="readonly",
                      font=F(10),bg="#eef6fb",fg=PRIMARY,
                      relief="solid",bd=1,readonlybackground="#eef6fb")
        pe.pack(side="left",fill="x",expand=True,ipady=6)
        tk.Button(rp,text="Clear",bg=WHITE,fg=DANGER,
                  relief="solid",bd=1,font=F(9),cursor="hand2",
                  padx=10,pady=4,command=self._clear_prev).pack(side="left",padx=(8,0))

        # MANUAL FETCH
        c4 = sec("MANUAL FETCH","🔍")
        rf=row(c4,"From Date & Time")
        self.dt_from=DateTimeField(rf,bg=WHITE,
                                   initial=datetime.now().replace(hour=0,minute=0,second=0))
        self.dt_from.pack(side="left")

        rt=row(c4,"To Date & Time")
        self.dt_to=DateTimeField(rt,bg=WHITE,
                                 initial=datetime.now().replace(hour=23,minute=59,second=59))
        self.dt_to.pack(side="left")

        # Buttons
        bf=tk.Frame(inner,bg=BG,pady=24); bf.pack(fill="x",padx=24)
        self._btn(bf,"💾  Save & Schedule",PRIMARY,WHITE,self._save).pack(side="left",padx=(0,10))
        self._btn(bf,"⬇  Fetch Data","#6c5ce7",WHITE,self._manual_fetch).pack(side="left",padx=(0,10))
        self._btn(bf,"▶  Sync Now",SUCCESS,WHITE,self._sync_now).pack(side="left",padx=(0,10))
        self._btn(bf,"🗑  Remove Task",DANGER,WHITE,self._remove_task).pack(side="left")

    def _btn(self,p,text,bg,fg,cmd):
        b=tk.Button(p,text=text,bg=bg,fg=fg,relief="flat",font=F(10,"bold"),
                    cursor="hand2",padx=16,pady=8,
                    activebackground=bg,activeforeground=fg,command=cmd)
        b.bind("<Enter>",lambda e,_b=bg:b.config(bg=self._dk(_b)))
        b.bind("<Leave>",lambda e,_b=bg:b.config(bg=_b))
        return b

    def _dk(self,h):
        try:
            r=max(0,int(h[1:3],16)-20); g=max(0,int(h[3:5],16)-20); b=max(0,int(h[5:7],16)-20)
            return f"#{r:02x}{g:02x}{b:02x}"
        except: return h

    def _hint(self):
        hints={"minute":"e.g. 5 = every 5 minutes","hour":"e.g. 1 = every 1 hour",
               "day":"e.g. 1 = every day","month":"e.g. 1 = every month"}
        self.hint_lbl.config(text=hints.get(self.v_stype.get(),""))

    def _build_logs(self,parent):
        top=tk.Frame(parent,bg=BG,pady=10); top.pack(fill="x",padx=16)
        tk.Label(top,text="Activity Logs",bg=BG,fg=TEXT,font=F(11,"bold")).pack(side="left")
        self._btn(top,"🗑  Clear",DANGER,WHITE,self._clear_logs).pack(side="right")
        self.log_box=scrolledtext.ScrolledText(parent,state="disabled",
                                               bg=WHITE,fg=TEXT,font=("Consolas",9),
                                               relief="flat",wrap="word",padx=12,pady=8)
        self.log_box.pack(fill="both",expand=True,padx=16,pady=(0,16))
        if os.path.exists(LOG_FILE):
            with open(LOG_FILE, encoding="utf-8", errors="ignore") as f: txt=f.read()
            self.log_box.config(state="normal")
            self.log_box.insert("end",txt)
            self.log_box.see("end")
            self.log_box.config(state="disabled")

    def _load_fields(self):
        self.v_folder.set(self.cfg.get("folder_path",""))
        self.v_file.set(self.cfg.get("file_name","eTimeTrackLite1.mdb"))
        self.v_url.set(self.cfg.get("url",""))
        self.v_user.set(self.cfg.get("username",""))
        self.v_pass.set(self.cfg.get("password",""))
        self.v_stype.set(self.cfg.get("schedule_type","hour"))
        self.v_sval.set(self.cfg.get("schedule_value","1"))
        prev=self.cfg.get("prev_synced_date","")
        self.v_prev.set(prev)
        if prev: self.dt_from.set_dt(prev)

    def _collect(self):
        self.cfg.update({"folder_path":self.v_folder.get().strip(),
                         "file_name":self.v_file.get().strip(),
                         "url":self.v_url.get().strip(),
                         "username":self.v_user.get().strip(),
                         "password":self.v_pass.get().strip(),
                         "schedule_type":self.v_stype.get(),
                         "schedule_value":self.v_sval.get().strip()})

    def _save(self):
        self._collect(); save_config(self.cfg)
        ok = register_task(self.cfg["schedule_type"],self.cfg["schedule_value"],self.log_box)
        write_log("Settings saved",self.log_box)
        msg = ("Settings saved!\n\n✅ Windows Task Scheduler updated.\n"
               "Sync runs automatically even after system restart."
               if ok else
               "Settings saved!\n\n⚠ Task Scheduler failed.\nTry running as Administrator.")
        (messagebox.showinfo if ok else messagebox.showwarning)("Saved",msg)

    def _clear_prev(self):
        self.cfg["prev_synced_date"]=""; self.v_prev.set(""); save_config(self.cfg)
        write_log("Prev synced date cleared — next sync sends ALL data",self.log_box)

    def _clear_logs(self):
        open(LOG_FILE,"w", encoding="utf-8").close()
        self.log_box.config(state="normal"); self.log_box.delete("1.0","end")
        self.log_box.config(state="disabled")

    def _remove_task(self):
        remove_task(self.log_box)
        messagebox.showinfo("Done","Scheduled task removed.")

    def _set_status(self,text,color=MUTED):
        self.status_lbl.config(text=text,fg=color)
        self.root.update_idletasks()

    def _manual_fetch(self):
        self._collect()
        f=self.dt_from.get_dt(); t=self.dt_to.get_dt()
        if not f or not t: messagebox.showerror("Error","Invalid date/time!"); return
        threading.Thread(target=self._run_fetch,args=(f,t,False),daemon=True).start()

    def _sync_now(self):
        threading.Thread(target=self._run_scheduled,daemon=True).start()

    def _run_fetch(self,from_dt,to_dt,update_prev):
        label="SCHEDULED SYNC" if update_prev else "MANUAL FETCH"
        self._set_status(f"⬤  {label}...","#f8a100")
        write_log(f"=== {label} ===",self.log_box)
        write_log(f"From: {from_dt or 'ALL'}  To: {to_dt}",self.log_box)
        mdb=os.path.join(self.cfg["folder_path"],self.cfg["file_name"])
        punches=fetch_punches(mdb,from_dt,to_dt,self.log_box)
        write_log(f"Records: {len(punches)}",self.log_box)
        if punches:
            url=self.cfg.get("url","")
            if url:
                ok=push_data(url,self.cfg.get("username",""),self.cfg.get("password",""),punches,self.log_box)
                if ok and update_prev: self._update_prev()
            else:
                write_log(f"No URL. Sample: {punches[:2]}",self.log_box)
                if update_prev: self._update_prev()
        else:
            write_log("No records",self.log_box)
            if update_prev: self._update_prev()
        write_log("Done.",self.log_box)
        self._set_status("⬤  Idle","#7f8c8d")

    def _update_prev(self):
        s = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.cfg["prev_synced_date"] = s
        save_config(self.cfg)
        # Must update UI from main thread
        self.root.after(0, lambda: self.v_prev.set(s))
        write_log(f"✅ prev_synced_date → {s}", self.log_box)

    def _run_scheduled(self):
        # Read prev_synced_date directly from saved config file (not from UI)
        # to avoid race condition when _save() is called before thread starts
        fresh_cfg = load_config()
        prev    = fresh_cfg.get("prev_synced_date", "")
        from_dt = None
        to_dt   = datetime.now()
        if prev:
            try: from_dt = datetime.strptime(prev, "%Y-%m-%d %H:%M:%S")
            except: pass
        self._run_fetch(from_dt, to_dt, update_prev=True)


if __name__=="__main__":
    if "--sync" in sys.argv:
        cfg=load_config(); prev=cfg.get("prev_synced_date","")
        from_dt=None; to_dt=datetime.now()
        if prev:
            try: from_dt=datetime.strptime(prev,"%Y-%m-%d %H:%M:%S")
            except: pass
        write_log("=== AUTO SYNC (Task Scheduler) ===")
        mdb=os.path.join(cfg["folder_path"],cfg["file_name"])
        punches=fetch_punches(mdb,from_dt,to_dt)
        write_log(f"Records: {len(punches)}")
        if punches and cfg.get("url"):
            ok=push_data(cfg["url"],cfg.get("username",""),cfg.get("password",""),punches)
            if ok:
                s=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                cfg["prev_synced_date"]=s; save_config(cfg)
                write_log(f"✅ Done. prev_synced_date → {s}")
        sys.exit(0)

    root=tk.Tk()
    App(root)
    root.mainloop()
