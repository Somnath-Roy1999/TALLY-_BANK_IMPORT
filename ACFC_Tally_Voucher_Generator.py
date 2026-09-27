import tkinter as tk
from tkinter import ttk,filedialog,messagebox
from pathlib import Path
from datetime import datetime
import openpyxl,xml.etree.ElementTree as ET,json,traceback

APP="ACFC Tally Voucher Generator"; COMPANY="ACFC E SERVICES INDIA PRIVATE LIMITED"
REQ=["Date","Narration","Withdrawal Amt.","Deposit Amt.","ledger type","Ledger Name","Journal Ledger","CONTRA BANK", "Payment bank", "Contra Credit bank", "Contra Debit Bank","Journal Payment Bank","Receipt Bank"]

def s(x): return str(x or "").strip()
def k(x): return s(x).casefold()
def amt(x):
    if x in (None,""): return 0.0
    return float(str(x).replace(",","").strip())
def date(x):
    if isinstance(x,datetime): return x.strftime("%Y%m%d")
    for f in ("%d/%m/%y","%d/%m/%Y","%Y-%m-%d","%d-%m-%Y","%m/%d/%Y"):
        try:return datetime.strptime(str(x).strip(),f).strftime("%Y%m%d")
        except:pass
    if isinstance(x,(int,float)):
        from openpyxl.utils.datetime import from_excel
        return from_excel(x).strftime("%Y%m%d")
    raise ValueError("Invalid date: "+repr(x))
def add(p,t,v=None,**a):
    e=ET.SubElement(p,t,a)
    if v is not None:e.text=str(v)
    return e

def master(path):
    w=openpyxl.load_workbook(path,data_only=True); ws=w["MASTER"] if "MASTER" in w.sheetnames else w.active
    h={k(ws.cell(1,c).value):c for c in range(1,ws.max_column+1)}
    if "ledger name" not in h or "under" not in h: raise ValueError("Master must contain Ledger Name and Under columns.")
    m={}
    for r in range(2,ws.max_row+1):
        n=s(ws.cell(r,h["ledger name"]).value); u=s(ws.cell(r,h["under"]).value)
        if n:m[k(n)]=(n,u or "Sundry Creditors")
    return m
def rows(path):
    w=openpyxl.load_workbook(path,data_only=True); ws=w.active
    h={k(ws.cell(1,c).value):c for c in range(1,ws.max_column+1)}
    miss=[x for x in REQ if k(x) not in h]
    if miss:raise ValueError("Bank Excel missing: "+", ".join(miss))
    out=[]
    for r in range(2,ws.max_row+1):
        z={x:ws.cell(r,h[k(x)]).value for x in REQ}
        if any(v not in (None,"") for v in z.values()):out.append(z)
    return out
def build(rs,m,auto=True):
    vs=[]; missing=[]; cnt={"Journal":0,"Payment":0,"Contra":0,"Receipt":0}; total=0
    q={x:0 for x in cnt}
    def ensure(n,u):
        n=s(n)
        if n and k(n) not in m:
            missing.append((n,u))
            if auto:m[k(n)]=(n,u)
    for i,r in enumerate(rs,2):
        t=k(r["ledger type"]); led=s(r["Ledger Name"]); d=date(r["Date"]); nar=s(r["Narration"]); w=amt(r["Withdrawal Amt."]); dep=amt(r["Deposit Amt."])
        if not led:raise ValueError(f"Row {i}: Ledger Name blank.")
        if t=="journal":
            a=w or dep; jl=s(r["Journal Ledger"]); bank=s(r["Journal Payment Bank"])
            if not jl or not bank:raise ValueError(f"Row {i}: Journal needs Journal Ledger and Journal Payment Bank.")
            ensure(jl,"Indirect Expenses");ensure(led,"Sundry Creditors");ensure(bank,"Bank Accounts")
            if a<=0:raise ValueError(f"Row {i}: Journal amount invalid.")
            q["Journal"]+=1;q["Payment"]+=1;cnt["Journal"]+=1;cnt["Payment"]+=1;total+=a
            vs += [(d,"Journal",f"J-{q['Journal']:04d}",nar,[(jl,-a,1),(led,a,0)]),(d,"Payment",f"P-{q['Payment']:04d}",nar,[(led,-a,1),(bank,a,0)])]
        elif t=="contra":
            a=dep or w; bank=s(r["CONTRA BANK", "Payment bank", "Contra Credit bank", "Contra Debit Bank"])
            if not bank:raise ValueError(f"Row {i}: Contra needs CONTRA BANK.")
            ensure(led,"Bank Accounts");ensure(bank,"Bank Accounts")
            if a<=0:raise ValueError(f"Row {i}: Contra amount invalid.")
            q["Contra"]+=1;cnt["Contra"]+=1;total+=a;vs.append((d,"Contra",f"C-{q['Contra']:04d}",nar,[(bank,-a,1),(led,a,0)]))
        elif t=="receipt":
            a=dep;bank=s(r["Receipt Bank"])
            if not bank:raise ValueError(f"Row {i}: Receipt needs Receipt Bank.")
            ensure(bank,"Bank Accounts");ensure(led,"Sundry Creditors")
            if a<=0:raise ValueError(f"Row {i}: Receipt amount invalid.")
            q["Receipt"]+=1;cnt["Receipt"]+=1;total+=a;vs.append((d,"Receipt",f"R-{q['Receipt']:04d}",nar,[(bank,-a,1),(led,a,0)]))
        else:raise ValueError(f"Row {i}: Unsupported ledger type {r['ledger type']!r}.")
    return vs,cnt,total,list(dict.fromkeys(missing))
def voucher_xml(vs,path):
    e=ET.Element("ENVELOPE");h=ET.SubElement(e,"HEADER")
    for t,v in [("VERSION","1"),("TALLYREQUEST","Import"),("TYPE","Data"),("ID","Vouchers")]:add(h,t,v)
    b=ET.SubElement(e,"BODY");d=ET.SubElement(b,"DESC");sv=ET.SubElement(d,"STATICVARIABLES");add(sv,"SVCURRENTCOMPANY",COMPANY);add(sv,"SVEXPORTFORMAT","$$SysName:XML");data=ET.SubElement(b,"DATA")
    for dt,typ,no,nar,ls in vs:
        v=ET.SubElement(ET.SubElement(data,"TALLYMESSAGE"),"VOUCHER",{"VCHTYPE":typ,"ACTION":"Create","OBJVIEW":"Accounting Voucher View"})
        for t,x in [("DATE",dt),("EFFECTIVEDATE",dt),("VOUCHERTYPENAME",typ),("VOUCHERNUMBER",no),("REFERENCE",no),("VCHENTRYMODE","Accounting Voucher"),("ISINVOICE","No"),("PERSISTEDVIEW","Accounting Voucher View"),("NARRATION",nar)]:add(v,t,x)
        for j,(ln,a,dr) in enumerate(ls):
            z=ET.SubElement(v,"ALLLEDGERENTRIES.LIST");add(z,"LEDGERNAME",ln);add(z,"ISDEEMEDPOSITIVE","Yes" if dr else "No");add(z,"ISLASTDEEMEDPOSITIVE","Yes" if j==len(ls)-1 else "No");add(z,"AMOUNT",f"{a:.2f}")
    ET.ElementTree(e).write(path,encoding="utf-8",xml_declaration=True)
def master_xml(m,path):
    e=ET.Element("ENVELOPE");h=ET.SubElement(e,"HEADER")
    for t,v in [("VERSION","1"),("TALLYREQUEST","Import"),("TYPE","Data"),("ID","All Masters")]:add(h,t,v)
    b=ET.SubElement(e,"BODY");d=ET.SubElement(b,"DESC");sv=ET.SubElement(d,"STATICVARIABLES");add(sv,"SVCURRENTCOMPANY",COMPANY);add(sv,"SVEXPORTFORMAT","$$SysName:XML");data=ET.SubElement(b,"DATA")
    for n,u in sorted(m.values(),key=lambda x:k(x[0])):
        l=ET.SubElement(ET.SubElement(data,"TALLYMESSAGE"),"LEDGER",{"NAME":n,"ACTION":"Create"});add(l,"NAME",n);add(l,"PARENT",u);add(l,"ISBILLWISEON","Yes" if k(u)=="sundry creditors" else "No");add(l,"AFFECTSSTOCK","No")
        if k(u)=="bank accounts":add(l,"ISBANKINGENABLED","Yes")
    ET.ElementTree(e).write(path,encoding="utf-8",xml_declaration=True)
def review(vs,path):
    w=openpyxl.Workbook();ws=w.active;ws.title="Voucher Review";ws.append(["Date","Voucher Type","Voucher No","Narration","Ledger Name","Under","DR","CR"])
    for dt,t,no,nar,ls in vs:
        for ln,a,dr in ls:ws.append([datetime.strptime(dt,"%Y%m%d").date(),t,no,nar,ln,"",abs(a) if dr else "",abs(a) if not dr else ""])
    ws.freeze_panes="A2";ws.auto_filter.ref=ws.dimensions
    for c in ws[1]:c.font=openpyxl.styles.Font(bold=True)
    for c,wid in zip("ABCDEFGH",[13,14,14,60,32,24,16,16]):ws.column_dimensions[c].width=wid
    w.save(path)
BUILTIN_MASTER = [("TIS EXPENSES","Indirect Expenses"),("BISWA BARAN DUBE","Sundry Creditors"),("MD PERSONAL A/C","Bank Accounts"),("HDFC BANK 3231","Bank Accounts"),("ACFC E SERVICES INDIA PRIVATE LIMITED","Sundry Creditors")]
def builtin_master(): return {k(n):(n,u) for n,u in BUILTIN_MASTER}
def generate(bank,out,auto=True):
    m=builtin_master();rs=rows(bank);vs,cnt,total,missing=build(rs,m,auto);o=Path(out);o.mkdir(parents=True,exist_ok=True);stem=Path(bank).stem.replace(" ","_")
    a=o/f"{stem}_Voucher_Import.xml";b=o/f"{stem}_Master_Import.xml";c=o/f"{stem}_Voucher_Review.xlsx";voucher_xml(vs,a);master_xml(m,b);review(vs,c)
    return len(rs),len(vs),cnt,total,missing,[a,b,c]

class App:
    def __init__(self,r):
        self.r=r;r.title(APP);r.geometry("900x620")
        self.bank=tk.StringVar();self.out=tk.StringVar();self.auto=tk.BooleanVar(value=True)
        ttk.Label(r,text=APP,font=("Segoe UI",20,"bold")).pack(pady=15);f=ttk.Frame(r);f.pack(fill="x",padx=25)
        self.row(f,0,"Bank Statement",self.bank,self.pb);self.row(f,1,"Output Folder",self.out,self.po)
        ttk.Checkbutton(r,text="Auto-create missing ledgers",variable=self.auto).pack(anchor="w",padx=25,pady=8)
        q=ttk.Frame(r);q.pack(pady=5);ttk.Button(q,text="CHECK DATA",command=self.check).pack(side="left",padx=5);ttk.Button(q,text="CREATE XML",command=self.create).pack(side="left",padx=5)
        self.log=tk.Text(r,font=("Consolas",10),state="disabled");self.log.pack(fill="both",expand=True,padx=25,pady=15)
    def row(self,f,i,t,v,cmd):
        ttk.Label(f,text=t,width=18).grid(row=i,column=0,sticky="w",pady=8);ttk.Entry(f,textvariable=v).grid(row=i,column=1,sticky="ew",padx=6);ttk.Button(f,text="Browse",command=cmd).grid(row=i,column=2);f.columnconfigure(1,weight=1)
    def pb(self): 
        p=filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xlsm")]);self.bank.set(p) if p else None
    def pm(self):
        p=filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xlsm")]);self.mp.set(p) if p else None
    def po(self):
        p=filedialog.askdirectory();self.out.set(p) if p else None
    def write(self,x):
        self.log.config(state="normal");self.log.insert("end",x+"\n");self.log.see("end");self.log.config(state="disabled")
    def check(self):
        try:
            rs=rows(self.bank.get());m=builtin_master();vs,c,t,mi=build(rs,m,False);self.write(f"Transactions: {len(rs)} | Vouchers: {len(vs)}");self.write(f"Journal {c['Journal']} | Payment {c['Payment']} | Contra {c['Contra']} | Receipt {c['Receipt']}");self.write(f"Total: ₹{t:,.2f} | Missing ledgers: {len({k(x[0]) for x in mi})}");messagebox.showinfo("Check",f"{len(rs)} transactions\n{len(vs)} vouchers\nMissing ledgers: {len({k(x[0]) for x in mi})}\nTotal: ₹{t:,.2f}")
        except Exception as e:messagebox.showerror("Error",str(e))
    def create(self):
        try:
            n,v,c,t,mi,fs=generate(self.bank.get(),self.out.get(),self.auto.get());self.write(f"Created {v} vouchers from {n} transactions | ₹{t:,.2f}");[self.write("→ "+str(x)) for x in fs];messagebox.showinfo("Completed",f"Created successfully!\n\nTransactions: {n}\nVouchers: {v}\nTotal: ₹{t:,.2f}")
        except Exception as e:self.write(traceback.format_exc());messagebox.showerror("Error",str(e))
tk.Tk.report_callback_exception=lambda *a:None
if __name__=="__main__":
 r=tk.Tk();App(r);r.mainloop()
