import React, { useState, useMemo, useEffect, useCallback } from "react";

// ─── SHEETJS (Excel export) ───────────────────────────────────────────────────
if (typeof window !== "undefined" && !window._xlsxLoaded) {
  window._xlsxLoaded = true;
  const s = document.createElement("script");
  s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
  document.head.appendChild(s);
}

// ─── SUPABASE ─────────────────────────────────────────────────────────────────
const SUPA_URL = "https://mmfhgdrnjobdpcsydssd.supabase.co";
const SUPA_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im1tZmhnZHJuam9iZHBjc3lkc3NkIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzMwOTMzNzYsImV4cCI6MjA4ODY2OTM3Nn0.sXqtcm65zzCzrDFGwN_BTAHD8UecEol93s2IQgIk3-Y";

async function sbFetch(path, opts = {}) {
  const session = JSON.parse(localStorage.getItem("sb_session") || "null");
  const headers = {
    "apikey": SUPA_KEY,
    "Content-Type": "application/json",
    "Prefer": "return=representation",
    ...(session?.access_token ? { "Authorization": `Bearer ${session.access_token}` } : {}),
    ...(opts.headers || {}),
  };
  const res = await fetch(`${SUPA_URL}${path}`, { ...opts, headers });
  if (res.status === 401) {
    // Sessão expirada — limpa e força logout
    localStorage.removeItem("sb_session");
    window.dispatchEvent(new Event("sb_session_expired"));
    throw new Error("Sessão expirada");
  }
  if (!res.ok) {
    const err = await res.text();
    throw new Error(err);
  }
  const text = await res.text();
  return text ? JSON.parse(text) : [];
}

async function sbLogin(email, password) {
  const res = await fetch(`${SUPA_URL}/auth/v1/token?grant_type=password`, {
    method: "POST",
    headers: { "apikey": SUPA_KEY, "Content-Type": "application/json" },
    body: JSON.stringify({ email, password }),
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error_description || "Erro ao fazer login");
  localStorage.setItem("sb_session", JSON.stringify(data));
  return data;
}

async function sbLogout() {
  const session = JSON.parse(localStorage.getItem("sb_session") || "null");
  if (session?.access_token) {
    await fetch(`${SUPA_URL}/auth/v1/logout`, {
      method: "POST",
      headers: { "apikey": SUPA_KEY, "Authorization": `Bearer ${session.access_token}` },
    }).catch(() => {});
  }
  localStorage.removeItem("sb_session");
}

function getSession() {
  return JSON.parse(localStorage.getItem("sb_session") || "null");
}

// Converter nota do banco para formato do app
function dbToNota(r) {
  return {
    id: r.id,
    fornecedor: r.fornecedor,
    fornecedorId: r.fornecedor_id,
    empresa: r.empresa,
    numero: r.numero,
    emissao: r.emissao,
    valor: parseFloat(r.valor),
    boletosRecebidos: r.boletos_recebidos,
    parcelas: r.parcelas,
    vencimentos: r.vencimentos || [],
    parcelasPagas: r.parcelas_pagas || [],
    observacao: r.observacao || "",
  };
}

// Converter nota do app para formato do banco
function notaToDB(n) {
  return {
    fornecedor: n.fornecedor,
    fornecedor_id: n.fornecedorId,
    empresa: n.empresa,
    numero: n.numero,
    emissao: n.emissao,
    valor: n.valor,
    boletos_recebidos: n.boletosRecebidos,
    parcelas: n.parcelas,
    vencimentos: n.vencimentos,
    parcelas_pagas: n.parcelasPagas || [],
    observacao: n.observacao || "",
  };
}

function dbToFornecedor(r) {
  return { id: r.id, nome: r.nome };
}

function dbToTarefa(r) {
  return {
    id: r.id,
    titulo: r.titulo,
    descricao: r.descricao || "",
    recorrencia: r.recorrencia,
    data: r.data,
    diaSemana: r.dia_semana,
    diaMes: r.dia_mes,
    concluidoEm: r.concluido_em,
  };
}

function tarefaToDB(t) {
  return {
    titulo: t.titulo,
    descricao: t.descricao || "",
    recorrencia: t.recorrencia,
    data: t.data,
    dia_semana: t.diaSemana,
    dia_mes: t.diaMes,
    concluido_em: t.concluidoEm,
  };
}


// ─── DESIGN TOKENS ────────────────────────────────────────────────────────────
const T = {
  font:  "'Inter', sans-serif",
  mono:  "'IBM Plex Mono', monospace",
  bg:    "#eef4f9",
  surface: "#ffffff",
  border: "#d6e6f2",
  borderStrong: "#5B89A6",
  text:   "#011526",
  textSub:"#5B89A6",
  textMuted:"#6c757d",
  primary:"#1A5173",
  primaryLight:"#C4DDF2",
  shadow: "0 1px 3px rgba(1,21,38,.06), 0 1px 2px rgba(1,21,38,.04)",
  shadowMd:"0 4px 16px rgba(1,21,38,.10)",
  radius: "10px",
  radiusSm:"7px",
  danger: "#F24E29",
  dangerLight:"#fdecea",
};

// ─── CONSTANTS ────────────────────────────────────────────────────────────────
const EMPRESAS = ["Kowalsky", "Evaldo", "Superfrio"];
const USER = { name: "Lukas", email: "lukaskowaa@gmail.com", initials: "L" };

const fmt   = v => Number(v).toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
const fdate = s => { if (!s) return "-"; const [y,m,d] = s.split("-"); return `${d}/${m}/${y}`; };
const fnome = s => s ? s.replace(/^FORN\d+ - /, "") : s;
const today = () => new Date().toISOString().split("T")[0];
const addMonths = (base, n) => { const d = new Date(base); d.setMonth(d.getMonth()+n); return d.toISOString().split("T")[0]; };
const tomorrow  = () => { const d = new Date(); d.setDate(d.getDate()+1); return d.toISOString().split("T")[0]; };
const endOfWeek = () => { const d = new Date(); d.setDate(d.getDate()+(6-d.getDay())); return d.toISOString().split("T")[0]; };

const S = {
  label: { display:"block", fontSize:"12.5px", fontWeight:600, color:T.textSub, marginBottom:"5px", letterSpacing:".01em" },
  input: { width:"100%", boxSizing:"border-box", padding:"9px 13px", borderRadius:T.radiusSm, border:`1.5px solid ${T.border}`, fontSize:"13.5px", color:T.text, outline:"none", background:T.surface, fontFamily:T.font, transition:"border-color .15s" },
  error: { fontSize:"11.5px", color:"#F24E29", marginTop:"3px", display:"block" },
  cbtn:  { width:"32px", height:"32px", borderRadius:T.radiusSm, border:`1.5px solid ${T.border}`, background:T.bg, fontSize:"17px", cursor:"pointer", color:T.textSub, display:"flex", alignItems:"center", justifyContent:"center", fontWeight:700 },
};
const Red = () => <span style={{color:"#F24E29"}}>*</span>;

// ═══════════════════════════════════════════════════════════════════════════
//  TOAST
// ═══════════════════════════════════════════════════════════════════════════
function ToastContainer({ toasts }) {
  return (
    <div style={{position:"fixed",bottom:"28px",left:"50%",transform:"translateX(-50%)",zIndex:200,display:"flex",flexDirection:"column",gap:"8px",alignItems:"center",pointerEvents:"none"}}>
      {toasts.map(t=>(
        <div key={t.id} style={{
          background: t.type==="error"?"#1e293b": t.type==="warning"?"#78350f":"#064e3b",
          color:"#fff",padding:"11px 20px",borderRadius:"10px",fontSize:"13.5px",fontWeight:600,
          boxShadow:"0 8px 24px rgba(0,0,0,0.18)",display:"flex",alignItems:"center",gap:"10px",
          fontFamily:T.font,whiteSpace:"nowrap",
          animation:"toastIn .25s ease",
        }}>
          <span style={{fontSize:"16px"}}>{t.type==="error"?"❌":t.type==="warning"?"⚠️":"✅"}</span>
          {t.msg}
        </div>
      ))}
      <style>{`@keyframes toastIn{from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:translateY(0)}}`}</style>
    </div>
  );
}

function useToast() {
  const [toasts, setToasts] = React.useState([]);
  function toast(msg, type="success") {
    const id = Date.now();
    setToasts(ts=>[...ts,{id,msg,type}]);
    setTimeout(()=>setToasts(ts=>ts.filter(t=>t.id!==id)), 2800);
  }
  return { toasts, toast };
}

// ═══════════════════════════════════════════════════════════════════════════
//  MODAL
// ═══════════════════════════════════════════════════════════════════════════
function NFModal({ onClose, onSave, editData, fornecedores, onNovoFornecedor, onFornecedorCriado }) {
  const blank = { fornecedor:"", fornecedorId:null, empresa:"", numero:"", emissao:today(), valor:"", boletosRecebidos:null, parcelas:1, vencimentos:[today()], status:"pendente", observacao:"" };
  // Quando um fornecedor novo é criado externamente, auto-seleciona
  useEffect(() => {
    if (onFornecedorCriado && typeof onFornecedorCriado === "function") return;
  }, []);
  const [form, setForm] = useState(editData || blank);
  const [errors, setErrors] = useState({});
  const [fornQuery, setFornQuery] = useState(editData?.fornecedor || "");
  const [showDropdown, setShowDropdown] = useState(false);
  const [temObs, setTemObs] = useState(!!(editData?.observacao));
  const semParcelas = !!(form.parcelasPagas && form.parcelas===1 && form.parcelasPagas.includes(form.vencimentos[0]));
  const set = (f,v) => { setForm(p=>({...p,[f]:v})); setErrors(p=>({...p,[f]:undefined})); };
  const EMPRESAS_LIST = ["Kowalsky","Evaldo","Superfrio"];

  // Ctrl+Enter salva de qualquer campo
  useEffect(() => {
    function onKey(e) {
      if (e.key === "Escape") { onClose(); return; }
      if ((e.ctrlKey||e.metaKey) && e.key === "Enter") { handleSave(); return; }
    }
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [form]);

  // Máscara de valor pt-BR
  function handleValorChange(e) {
    let raw = e.target.value.replace(/\D/g,"");
    if (!raw) { set("valor",""); return; }
    const num = parseInt(raw,10)/100;
    const formatted = num.toLocaleString("pt-BR",{minimumFractionDigits:2,maximumFractionDigits:2});
    set("valor", formatted);
  }
  function parseValor(v) {
    if (!v) return NaN;
    return parseFloat(v.toString().replace(/\./g,"").replace(",","."));
  }

  const fornSuggestions = useMemo(() => {
    if (!fornQuery.trim()) return fornecedores;
    const q = fornQuery.toLowerCase();
    return fornecedores.filter(f => f.nome.toLowerCase().includes(q));
  }, [fornQuery, fornecedores]);

  function selectFornecedor(f) {
    setFornQuery(f.nome);
    set("fornecedor", f.nome);
    set("fornecedorId", f.id);
    setShowDropdown(false);
  }

  function setParcelas(n) {
    const count = Math.max(1,Math.min(12,n));
    const base  = form.vencimentos[0]||today();
    const venc  = Array.from({length:count},(_,i)=>form.vencimentos[i]||addMonths(base,i));
    setForm(p=>({...p,parcelas:count,vencimentos:venc}));
  }
  function validate() {
    const e={};
    if(!form.fornecedor.trim()) e.fornecedor="Obrigatório";
    if(!form.empresa)           e.empresa="Selecione a empresa";
    if(!form.numero.trim())     e.numero="Obrigatório";
    if(!form.emissao||!form.emissao.match(/^\d{4}-\d{2}-\d{2}$/)) e.emissao="Data inválida (use DD/MM/AAAA)";
    if(!form.valor||isNaN(parseValor(form.valor))) e.valor="Valor inválido";
    if(form.boletosRecebidos===null) e.boletosRecebidos="Selecione uma opção";
    const isSemParcelas = !!(form.parcelasPagas && form.parcelas===1 && form.parcelasPagas.includes(form.vencimentos[0]));
    if (!isSemParcelas) {
      const allVencValid = form.vencimentos.every(v=>v&&v.match(/^\d{4}-\d{2}-\d{2}$/));
      if(!allVencValid) e.vencimentos="Preencha todas as datas no formato DD/MM/AAAA";
      else {
        const vencAnterior = form.vencimentos.some(v=>v && v < form.emissao);
        if(vencAnterior) e.vencimentos="Uma ou mais datas de vencimento são anteriores à emissão";
      }
    }
    setErrors(e); return !Object.keys(e).length;
  }
  function handleSave() {
    if(!validate()) return;
    onSave({...form, valor:parseValor(form.valor), id:editData?.id||Date.now()});
  }

  return (
    <div style={{position:"fixed",inset:0,background:"rgba(15,23,42,0.5)",backdropFilter:"blur(4px)",zIndex:60,display:"flex",alignItems:"center",justifyContent:"center",padding:"16px"}}>
      <div style={{background:T.surface,borderRadius:"16px",width:"100%",maxWidth:"510px",maxHeight:"92vh",overflowY:"auto",boxShadow:"0 24px 64px rgba(0,0,0,0.18)",fontFamily:T.font}}>
        <div style={{padding:"22px 26px 0",display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
          <div>
            <h2 style={{margin:0,fontSize:"18px",fontWeight:800,color:T.text}}>{editData?"Editar Nota Fiscal":"Nova Nota Fiscal"}</h2>
            <p style={{margin:"3px 0 0",fontSize:"12.5px",color:T.textSub}}>Preencha os dados da nota fiscal</p>
          </div>
          <button onClick={onClose} style={{border:"none",background:"none",cursor:"pointer",fontSize:"22px",color:T.textMuted}}>×</button>
        </div>
        <div style={{padding:"18px 26px 26px",display:"flex",flexDirection:"column",gap:"15px"}}>
          <div style={{position:"relative"}}>
            <label style={S.label}>Fornecedor <Red/></label>
            <input
              value={fornQuery}
              onChange={e=>{ setFornQuery(e.target.value); set("fornecedor",e.target.value); set("fornecedorId",null); setShowDropdown(true); }}
              onFocus={()=>setShowDropdown(true)}
              onBlur={()=>setTimeout(()=>setShowDropdown(false),200)}
              placeholder="Digite o código ou nome do fornecedor"
              style={{...S.input,borderColor:errors.fornecedor?"#F24E29":"#e2e8f0"}}
            />
            {errors.fornecedor&&<span style={S.error}>{errors.fornecedor}</span>}
            {showDropdown&&(
              <div style={{position:"absolute",top:"100%",left:0,right:0,background:T.surface,borderRadius:T.radiusSm,boxShadow:"0 8px 24px rgba(0,0,0,0.12)",border:`1px solid ${T.border}`,zIndex:99,maxHeight:"180px",overflowY:"auto",marginTop:"4px"}}>
                {fornSuggestions.length===0&&fornQuery.trim()===("")?(
                  <div style={{padding:"12px 14px",fontSize:"13px",color:T.textMuted}}>Nenhum fornecedor cadastrado</div>
                ):fornSuggestions.map(f=>(
                  <button key={f.id} onMouseDown={()=>selectFornecedor(f)} style={{width:"100%",padding:"10px 14px",border:"none",background:"none",cursor:"pointer",textAlign:"left",fontSize:"13.5px",color:T.text,fontFamily:"inherit",borderBottom:"1px solid #f1f5f9"}}
                    onMouseEnter={e=>e.currentTarget.style.background="#f8fafc"}
                    onMouseLeave={e=>e.currentTarget.style.background="none"}>
                    {f.nome}
                  </button>
                ))}
                <button onMouseDown={()=>{ setShowDropdown(false); onNovoFornecedor(fornQuery, (novoForn)=>{ setFornQuery(novoForn.nome); set("fornecedor", novoForn.nome); set("fornecedorId", novoForn.id); }); }} style={{width:"100%",padding:"10px 14px",border:"none",borderTop:"1px solid #e2e8f0",background:"none",cursor:"pointer",textAlign:"left",fontSize:"13.5px",color:"#1A5173",fontWeight:700,fontFamily:"inherit",display:"flex",alignItems:"center",gap:"6px"}}>
                  <span style={{fontSize:"16px"}}>+</span> Criar novo fornecedor
                </button>
              </div>
            )}
          </div>
          <div>
            <label style={S.label}>Empresa Destino <Red/></label>
            <div role="group" style={{display:"flex",gap:"8px"}}
              onKeyDown={e=>{
                const idx = EMPRESAS_LIST.indexOf(form.empresa);
                if(e.key==="ArrowRight"||e.key==="ArrowDown"){ e.preventDefault(); set("empresa",EMPRESAS_LIST[(idx+1)%EMPRESAS_LIST.length]); }
                if(e.key==="ArrowLeft"||e.key==="ArrowUp"){ e.preventDefault(); set("empresa",EMPRESAS_LIST[(idx-1+EMPRESAS_LIST.length)%EMPRESAS_LIST.length]); }
              }}>
              {EMPRESAS.map(e=>(
                <button key={e} tabIndex={0} onClick={()=>set("empresa",e)}
                  style={{flex:1,padding:"9px 10px",borderRadius:T.radiusSm,cursor:"pointer",fontWeight:600,fontSize:"13px",transition:"all .15s",border:form.empresa===e?"2px solid #1A5173":"2px solid #e2e8f0",background:form.empresa===e?"#C4DDF2":"#f8fafc",color:form.empresa===e?"#1A5173":"#475569"}}>{e}</button>
              ))}
            </div>
            {errors.empresa&&<span style={S.error}>{errors.empresa}</span>}
          </div>
          <div>
            <label style={S.label}>Número da Nota Fiscal <Red/></label>
            <input value={form.numero} onChange={e=>set("numero",e.target.value)} placeholder="Ex: NF-001" style={{...S.input,borderColor:errors.numero?"#F24E29":"#e2e8f0"}}/>
            {errors.numero&&<span style={S.error}>{errors.numero}</span>}
          </div>
          <div>
            <label style={S.label}>Data de Emissão <Red/></label>
            <input
              type="text"
              inputMode="numeric"
              placeholder="DD/MM/AAAA"
              value={form.emissao ? form.emissao.split("-").reverse().join("/") : ""}
              onChange={e=>{
                let v = e.target.value.replace(/\D/g,"");
                if(v.length>8) v=v.slice(0,8);
                // Auto-preenche o ano atual quando digitar só dia/mês (4 dígitos)
                const anoAtual = String(new Date().getFullYear());
                if(v.length===4){
                  v = v + anoAtual;
                }
                let disp="";
                if(v.length<=2) disp=v;
                else if(v.length<=4) disp=v.slice(0,2)+"/"+v.slice(2);
                else disp=v.slice(0,2)+"/"+v.slice(2,4)+"/"+v.slice(4);
                if(v.length===8){
                  const iso=v.slice(4)+"-"+v.slice(2,4)+"-"+v.slice(0,2);
                  set("emissao",iso);
                } else {
                  set("emissao",disp);
                }
              }}
              onFocus={e=>e.target.select()}
              style={{...S.input,borderColor:errors.emissao?"#F24E29":"#e2e8f0"}}
            />
            {errors.emissao&&<span style={S.error}>{errors.emissao}</span>}
          </div>
          <div>
            <label style={S.label}>Valor total da Nota <Red/></label>
            <div style={{position:"relative"}}>
              <span style={{position:"absolute",left:"13px",top:"50%",transform:"translateY(-50%)",color:T.textSub,fontWeight:600,fontSize:"13px"}}>R$</span>
              <input
                value={form.valor}
                onChange={handleValorChange}
                placeholder="0,00"
                inputMode="numeric"
                style={{...S.input,paddingLeft:"38px",borderColor:errors.valor?"#F24E29":"#e2e8f0"}}
              />
            </div>
            {errors.valor&&<span style={S.error}>{errors.valor}</span>}
            {form.parcelas>1 && form.valor && !isNaN(parseValor(form.valor)) && (
              <span style={{fontSize:"11.5px",color:T.textSub,marginTop:"4px",display:"block"}}>
                ≈ {fmt(parseValor(form.valor)/form.parcelas)} por parcela
              </span>
            )}
          </div>
          <div>
            <label style={S.label}>Boletos físicos recebidos? <Red/></label>
            <div role="group" style={{display:"flex",gap:"8px"}}
              onKeyDown={e=>{
                if(e.key==="ArrowLeft"||e.key==="ArrowRight"||e.key==="ArrowUp"||e.key==="ArrowDown"){
                  e.preventDefault();
                  set("boletosRecebidos", form.boletosRecebidos===true ? false : true);
                }
              }}>
              {[true,false].map(v=>(
                <button key={String(v)} tabIndex={0} onClick={()=>set("boletosRecebidos",v)}
                  style={{flex:1,padding:"9px",borderRadius:T.radiusSm,cursor:"pointer",fontWeight:600,fontSize:"13.5px",border:form.boletosRecebidos===v?"2px solid #1A5173":"2px solid #e2e8f0",background:form.boletosRecebidos===v?"#C4DDF2":"#f8fafc",color:form.boletosRecebidos===v?"#1A5173":"#475569"}}>{v?"Sim":"Não"}</button>
              ))}
            </div>
            {errors.boletosRecebidos&&<span style={S.error}>{errors.boletosRecebidos}</span>}
          </div>
          <div>
            <label style={S.label}>Quantidade de parcelas <Red/></label>
            <div style={{display:"flex",alignItems:"center",gap:"12px",flexWrap:"wrap"}}>
              <button onClick={()=>setParcelas(form.parcelas-1)} style={S.cbtn}>−</button>
              <input
                type="number" min="1" max="12" value={form.parcelas}
                onChange={e=>setParcelas(parseInt(e.target.value)||1)}
                onKeyDown={e=>{
                  if(e.key==="ArrowUp"){e.preventDefault();setParcelas(form.parcelas+1);}
                  if(e.key==="ArrowDown"){e.preventDefault();setParcelas(form.parcelas-1);}
                }}
                style={{width:"44px",textAlign:"center",fontSize:"16px",fontWeight:700,color:T.text,border:`1.5px solid ${T.border}`,borderRadius:T.radiusSm,padding:"4px 0",background:T.bg,fontFamily:T.font,outline:"none"}}
              />
              <button onClick={()=>setParcelas(form.parcelas+1)} style={S.cbtn}>+</button>
              <button type="button" onClick={()=>{
                const dataBase = form.emissao && form.emissao.match(/^\d{4}-\d{2}-\d{2}$/) ? form.emissao : today();
                setForm(p=>({...p,parcelas:1,vencimentos:[dataBase],parcelasPagas:[dataBase],boletosRecebidos:true}));
              }} style={{marginLeft:"8px",padding:"6px 14px",borderRadius:"20px",fontSize:"12px",fontWeight:700,cursor:"pointer",border:"none",
                background:semParcelas?"#059669":T.bg,
                color:semParcelas?"#fff":T.textSub,
                outline:semParcelas?"none":`1.5px solid ${T.border}`,
                transition:"all .12s"}}>
                ✓ Sem parcelas
              </button>
            </div>
            {semParcelas && (
              <span style={{fontSize:"12px",color:"#059669",marginTop:"4px",display:"block",fontWeight:600}}>✓ Nota marcada como quitada — nenhuma parcela a pagar</span>
            )}
          </div>
          {!semParcelas && <div>
            <label style={S.label}>Datas de Vencimento das Parcelas <Red/></label>
            <div style={{display:"flex",flexDirection:"column",gap:"7px"}}>
              {Array.from({length:form.parcelas},(_,i)=>(
                <div key={i} style={{display:"flex",alignItems:"center",gap:"10px"}}>
                  <span style={{fontSize:"12.5px",color:T.textSub,minWidth:"60px"}}>Parcela {i+1}</span>
                  <input
                    type="text" inputMode="numeric" placeholder="DD/MM/AAAA"
                    data-venc-idx={i}
                    value={form.vencimentos[i] && form.vencimentos[i].includes("-") ? form.vencimentos[i].split("-").reverse().join("/") : (form.vencimentos[i]||"")}
                    onChange={e=>{
                      let raw=e.target.value.replace(/\D/g,"");
                      if(raw.length>8) raw=raw.slice(0,8);
                      const anoAtual=String(new Date().getFullYear());
                      if(raw.length===4) raw=raw+anoAtual;
                      let disp="";
                      if(raw.length<=2) disp=raw;
                      else if(raw.length<=4) disp=raw.slice(0,2)+"/"+raw.slice(2);
                      else disp=raw.slice(0,2)+"/"+raw.slice(2,4)+"/"+raw.slice(4);
                      const arr=[...form.vencimentos];
                      if(raw.length===8){
                        arr[i]=raw.slice(4)+"-"+raw.slice(2,4)+"-"+raw.slice(0,2);
                      } else {
                        arr[i]=disp;
                      }
                      setForm(p=>({...p,vencimentos:arr}));
                      setErrors(p=>({...p,vencimentos:undefined}));
                    }}
                    onFocus={e=>{ e.target.select(); }}
                    onKeyDown={e=>{
                      if(e.key==="Tab" && !e.shiftKey && i < form.parcelas-1) {
                        e.preventDefault();
                        const next = document.querySelector(`[data-venc-idx="${i+1}"]`);
                        if(next) { next.focus(); next.select(); }
                      }
                    }}
                    style={{...S.input,flex:1,borderColor:(form.vencimentos[i]&&form.emissao&&form.vencimentos[i]<form.emissao)?"#F24E29":"#e2e8f0"}}
                  />
                </div>
              ))}
            </div>
            {errors.vencimentos&&<span style={S.error}>{errors.vencimentos}</span>}
          </div>}
          <div>
            <label style={S.label}>Observação</label>
            <div style={{display:"flex",gap:"6px",marginBottom:temObs?"10px":"0"}}>
              {[["nao","Não"],["sim","Sim"]].map(([v,l])=>(
                <button key={v} type="button"
                  onClick={()=>{ setTemObs(v==="sim"); if(v==="nao") set("observacao",""); }}
                  style={{padding:"6px 18px",borderRadius:"20px",fontSize:"12.5px",fontWeight:600,cursor:"pointer",border:"none",
                    background:(v==="sim"?temObs:!temObs)?T.primary:T.bg,
                    color:(v==="sim"?temObs:!temObs)?"#fff":T.textMuted,
                    outline:(v==="sim"?temObs:!temObs)?"none":`1.5px solid ${T.border}`,
                    transition:"all .12s"}}>
                  {l}
                </button>
              ))}
            </div>
            {temObs && (
              <textarea value={form.observacao} onChange={e=>set("observacao",e.target.value)}
                placeholder="Digite a observação..." rows={2} autoFocus
                style={{...S.input,resize:"vertical",minHeight:"56px"}}/>
            )}
          </div>

          <div style={{display:"flex",gap:"10px",marginTop:"4px"}}>
            <button onClick={onClose} style={{flex:1,padding:"11px",borderRadius:"9px",border:"1.5px solid #e2e8f0",background:T.bg,color:"#475569",fontWeight:600,fontSize:"13.5px",cursor:"pointer"}}>Cancelar</button>
            <button onClick={handleSave} style={{flex:2,padding:"11px",borderRadius:"9px",border:"none",background:"linear-gradient(135deg,#1A5173,#1A5173)",color:"#fff",fontWeight:700,fontSize:"13.5px",cursor:"pointer",boxShadow:"0 4px 14px rgba(37,99,235,.3)"}}>{editData?"Salvar alterações":"Criar Nota Fiscal"}</button>
          </div>
          <p style={{margin:"10px 0 0",textAlign:"center",fontSize:"11px",color:T.textMuted}}>
            <kbd style={{background:"#f1f5f9",border:"1px solid #e2e8f0",borderRadius:"4px",padding:"1px 5px",fontSize:"10px",fontFamily:T.mono}}>Ctrl+Enter</kbd> salva &nbsp;·&nbsp;
            <kbd style={{background:"#f1f5f9",border:"1px solid #e2e8f0",borderRadius:"4px",padding:"1px 5px",fontSize:"10px",fontFamily:T.mono}}>Esc</kbd> fecha &nbsp;·&nbsp;
            <kbd style={{background:"#f1f5f9",border:"1px solid #e2e8f0",borderRadius:"4px",padding:"1px 5px",fontSize:"10px",fontFamily:T.mono}}>← →</kbd> seleciona opções
          </p>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════
//  BAR CHART DIÁRIO
// ═══════════════════════════════════════════════════════════════════════════
function BarChartDiario({ notas }) {
  const now = new Date();
  const [mesOffset, setMesOffset] = React.useState(0); // 0 = atual, 1 = próximo
  const [tooltip, setTooltip] = React.useState(null);  // {day, value, parcelas, x, y}

  const ano  = now.getFullYear();
  const mes  = now.getMonth() + mesOffset;
  const d    = new Date(ano, mes, 1);
  const mesKey  = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;
  const mesNome = d.toLocaleString("pt-BR",{month:"long",year:"numeric"});
  const diasNoMes = new Date(d.getFullYear(), d.getMonth()+1, 0).getDate();
  const todayStr = today();

  // Calcular valor total de parcelas vencendo por dia
  const dados = Array.from({length:diasNoMes},(_,i)=>{
    const dia = i+1;
    const diaStr = `${mesKey}-${String(dia).padStart(2,"0")}`;
    let total = 0, parcelas = 0;
    notas.forEach(n=>{
      n.vencimentos.forEach(v=>{
        if(v === diaStr && parcelaStatusEfetivo(n,v)!=="pago") {
          total += n.valor/n.parcelas;
          parcelas++;
        }
      });
    });
    return { dia, diaStr, total, parcelas };
  });

  const maxVal = Math.max(...dados.map(d=>d.total), 1);
  const CHART_H = 140;

  // Formatar eixo Y
  function fmtY(v) {
    if(v>=1000000) return `${(v/1000000).toFixed(1)}M`;
    if(v>=1000)    return `${(v/1000).toFixed(0)}k`;
    return v.toFixed(0);
  }

  // Linhas do eixo Y (4 níveis)
  const yTicks = [0.25,0.5,0.75,1.0].map(p=>({ pct:p, val:maxVal*p }));

  return (
    <div style={{background:T.surface,borderRadius:T.radius,padding:"20px 22px",boxShadow:T.shadow,border:`1px solid ${T.border}`,fontFamily:T.font}}>
      {/* Header */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:"18px"}}>
        <div>
          <p style={{margin:0,fontSize:"13px",fontWeight:600,color:T.text}}>Vencimentos do Mês</p>
          <p style={{margin:"2px 0 0",fontSize:"12px",color:T.textMuted,textTransform:"capitalize"}}>{mesNome} — valor diário de parcelas</p>
        </div>
        <div style={{display:"flex",gap:"4px"}}>
          <button onClick={()=>setMesOffset(0)}
            style={{padding:"4px 12px",borderRadius:T.radiusSm,border:`1.5px solid ${mesOffset===0?T.primary:T.border}`,background:mesOffset===0?T.primaryLight:T.surface,color:mesOffset===0?T.primary:T.textMuted,fontSize:"12px",fontWeight:600,cursor:"pointer",fontFamily:T.font}}>
            Este mês
          </button>
          <button onClick={()=>setMesOffset(1)}
            style={{padding:"4px 12px",borderRadius:T.radiusSm,border:`1.5px solid ${mesOffset===1?T.primary:T.border}`,background:mesOffset===1?T.primaryLight:T.surface,color:mesOffset===1?T.primary:T.textMuted,fontSize:"12px",fontWeight:600,cursor:"pointer",fontFamily:T.font}}>
            Próximo mês
          </button>
        </div>
      </div>

      {/* Chart area */}
      <div data-chart style={{position:"relative"}}>
        {/* Eixo Y — linhas de referência */}
        <div style={{position:"absolute",top:0,left:0,right:0,height:CHART_H,pointerEvents:"none"}}>
          {yTicks.map(t=>(
            <div key={t.pct} style={{position:"absolute",bottom:`${t.pct*100}%`,left:0,right:0,display:"flex",alignItems:"center",gap:"6px"}}>
              <span style={{fontSize:"10px",color:T.textMuted,fontFamily:T.mono,minWidth:"32px",textAlign:"right",flexShrink:0}}>{fmtY(t.val)}</span>
              <div style={{flex:1,height:"1px",background:T.border}}/>
            </div>
          ))}
          {/* linha zero */}
          <div style={{position:"absolute",bottom:0,left:38,right:0,height:"1px",background:T.borderStrong}}/>
        </div>

        {/* Barras */}
        <div style={{marginLeft:"38px",height:CHART_H,display:"flex",alignItems:"flex-end",gap:"2px",paddingBottom:"0"}}>
          {dados.map(d=>{
            const isHoje = d.diaStr === todayStr;
            const hasVal = d.total > 0;
            const barH   = hasVal ? Math.max((d.total/maxVal)*CHART_H, 3) : 2;
            const barColor = isHoje ? T.danger : hasVal ? T.primary : T.border;
            return (
              <div key={d.dia} style={{flex:1,display:"flex",flexDirection:"column",justifyContent:"flex-end",height:CHART_H,position:"relative",cursor:hasVal?"pointer":"default"}}
                onMouseEnter={e=>{
                  if(!hasVal && !isHoje) return;
                  const rect = e.currentTarget.getBoundingClientRect();
                  const parent = e.currentTarget.closest('[data-chart]').getBoundingClientRect();
                  setTooltip({dia:d.dia,total:d.total,parcelas:d.parcelas,diaStr:d.diaStr,
                    x: rect.left - parent.left + rect.width/2,
                    y: rect.top - parent.top - 8
                  });
                }}
                onMouseLeave={()=>setTooltip(null)}>
                <div style={{
                  width:"100%", height:`${barH}px`,
                  background: barColor,
                  borderRadius:"3px 3px 0 0",
                  transition:"opacity .1s",
                  opacity: tooltip && tooltip.dia!==d.dia ? 0.6 : 1,
                }}/>
              </div>
            );
          })}
        </div>

        {/* Tooltip */}
        {tooltip && (
          <div style={{
            position:"absolute", pointerEvents:"none", zIndex:10,
            left: tooltip.x + 38, top: tooltip.y,
            transform:"translate(-50%,-100%)",
            background:T.text, color:"#fff",
            borderRadius:"6px", padding:"7px 11px",
            fontSize:"12px", fontFamily:T.font,
            boxShadow:T.shadowMd, whiteSpace:"nowrap",
          }}>
            <div style={{fontWeight:700,marginBottom:"2px"}}>{tooltip.diaStr.split("-").reverse().slice(0,2).join("/")}</div>
            <div style={{fontFamily:T.mono,fontSize:"13px",fontWeight:600}}>{fmt(tooltip.total)}</div>
            <div style={{color:"#adb5bd",fontSize:"11px",marginTop:"1px"}}>{tooltip.parcelas} parcela(s)</div>
          </div>
        )}

        {/* Eixo X — todos os dias */}
        <div style={{marginLeft:"38px",display:"flex",gap:"2px",marginTop:"5px"}}>
          {dados.map(d=>(
            <div key={d.dia} style={{flex:1,textAlign:"center",fontSize:"9px",color:d.diaStr===todayStr?T.danger:T.textMuted,fontWeight:d.diaStr===todayStr?700:400,fontFamily:T.mono}}>
              {d.dia}
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════
//  HOME
const PIE_COLORS = ["#1A5173","#F24E29","#5B89A6","#17b26a","#f79009","#7c3aed","#0ea5e9","#db2777","#64748b"];

function PieChartFornecedores({ notas }) {
  const [periodo, setPeriodo] = React.useState("ano");
  const [hover, setHover]     = React.useState(null);
  const td  = today();
  const mesKey = `${new Date().getFullYear()}-${String(new Date().getMonth()+1).padStart(2,"0")}`;
  const p3  = (() => { const d = new Date(); d.setMonth(d.getMonth()+3); return d.toISOString().split("T")[0]; })();

  const periodoFn = {
    mes:   v => v.startsWith(mesKey),
    trim:  v => v >= td && v <= p3,
    ano:   v => v.startsWith(new Date().getFullYear().toString()),
  }[periodo];

  // Agrupar por fornecedor — APENAS parcelas em aberto (não pagas)
  const mapa = {};
  notas.forEach(n => {
    const nome = fnome(n.fornecedor) || "Sem fornecedor";
    n.vencimentos.forEach(v => {
      if (!periodoFn(v)) return;
      if (parcelaStatusEfetivo(n, v) === "pago") return; // ignora pagas
      if (!mapa[nome]) mapa[nome] = 0;
      mapa[nome] += n.valor / n.parcelas;
    });
  });

  // Ordenar por valor desc, limitar 8 + "Outros"
  let itens = Object.entries(mapa)
    .map(([nome, valor]) => ({ nome, valor }))
    .sort((a, b) => b.valor - a.valor);

  if (itens.length > 8) {
    const outros = itens.slice(8).reduce((s, i) => s + i.valor, 0);
    itens = [...itens.slice(0, 8), ...(outros > 0 ? [{ nome: "Outros", valor: outros }] : [])];
  }

  const total = itens.reduce((s, i) => s + i.valor, 0);

  if (total === 0) return (
    <div style={{background:T.surface,borderRadius:T.radius,padding:"24px",boxShadow:T.shadow,border:`1px solid ${T.border}`,marginTop:"16px"}}>
      <p style={{margin:0,fontSize:"13px",fontWeight:600,color:T.text}}>Valor em Aberto por Fornecedor</p>
      <p style={{margin:"24px 0",textAlign:"center",color:T.textMuted,fontSize:"13px"}}>Nenhuma parcela em aberto no período</p>
    </div>
  );

  // Calcular fatias SVG (donut)
  const SIZE = 190, CX = SIZE/2, CY = SIZE/2, R = 80, RI = 46;
  let angle = -Math.PI / 2;
  const fatias = itens.map((it, idx) => {
    const pct   = it.valor / total;
    const sweep = pct * 2 * Math.PI;
    const a1 = angle, a2 = angle + sweep;
    const cos1 = Math.cos(a1), sin1 = Math.sin(a1);
    const cos2 = Math.cos(a2), sin2 = Math.sin(a2);
    const large = sweep > Math.PI ? 1 : 0;
    const path = [
      `M ${CX+RI*cos1} ${CY+RI*sin1}`,
      `L ${CX+R*cos1} ${CY+R*sin1}`,
      `A ${R} ${R} 0 ${large} 1 ${CX+R*cos2} ${CY+R*sin2}`,
      `L ${CX+RI*cos2} ${CY+RI*sin2}`,
      `A ${RI} ${RI} 0 ${large} 0 ${CX+RI*cos1} ${CY+RI*sin1}`,
      "Z"
    ].join(" ");
    angle = a2;
    return { ...it, path, color: PIE_COLORS[idx % PIE_COLORS.length], pct, idx };
  });

  const periodoLabel = { mes: "Este mês", trim: "Próximos 3 meses", ano: "Este ano" }[periodo];

  return (
    <div style={{background:T.surface,borderRadius:T.radius,padding:"20px 22px",boxShadow:T.shadow,border:`1px solid ${T.border}`,marginTop:"16px"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:"18px"}}>
        <div>
          <p style={{margin:0,fontSize:"13px",fontWeight:600,color:T.text}}>Valor em Aberto por Fornecedor</p>
          <p style={{margin:"2px 0 0",fontSize:"11.5px",color:T.textMuted}}>Apenas parcelas não quitadas — {periodoLabel.toLowerCase()}</p>
        </div>
        <div style={{display:"flex",gap:"4px"}}>
          {[["mes","Mês"],["trim","3 meses"],["ano","Ano"]].map(([k,l])=>(
            <button key={k} onClick={()=>setPeriodo(k)}
              style={{padding:"5px 11px",borderRadius:"20px",fontSize:"12px",fontWeight:600,cursor:"pointer",border:"none",
                background:periodo===k?T.primary:"transparent",
                color:periodo===k?"#fff":T.textSub,
                outline:periodo!==k?`1.5px solid ${T.border}`:"none"}}>
              {l}
            </button>
          ))}
        </div>
      </div>

      <div style={{display:"flex",alignItems:"center",gap:"32px",flexWrap:"wrap"}}>
        {/* Donut SVG */}
        <div style={{position:"relative",flexShrink:0}}>
          <svg width={SIZE} height={SIZE} style={{overflow:"visible"}}>
            {fatias.map((f,i)=>(
              <path key={f.nome} d={f.path} fill={f.color}
                opacity={hover!==null && hover!==i ? 0.5 : 1}
                style={{cursor:"pointer",transition:"all .15s",
                  transform: hover===i ? "scale(1.05)" : "scale(1)",
                  transformOrigin:`${CX}px ${CY}px`,
                  filter: hover===i ? "drop-shadow(0 2px 6px rgba(0,0,0,.2))" : "none"}}
                onMouseEnter={()=>setHover(i)}
                onMouseLeave={()=>setHover(null)}/>
            ))}
          </svg>
          {/* Centro */}
          <div style={{position:"absolute",top:"50%",left:"50%",transform:"translate(-50%,-50%)",textAlign:"center",pointerEvents:"none",width:"70px"}}>
            {hover!==null ? (
              <>
                <div style={{fontSize:"9.5px",color:T.textMuted,lineHeight:1.3,marginBottom:"2px"}}>{fatias[hover].nome}</div>
                <div style={{fontSize:"13px",fontWeight:800,color:fatias[hover].color}}>{Math.round(fatias[hover].pct*100)}%</div>
              </>
            ) : (
              <>
                <div style={{fontSize:"9.5px",color:T.textMuted}}>Total</div>
                <div style={{fontSize:"10.5px",fontWeight:800,color:T.text,fontFamily:T.mono}}>{fmt(total)}</div>
              </>
            )}
          </div>
        </div>

        {/* Legenda */}
        <div style={{flex:1,minWidth:"200px",display:"flex",flexDirection:"column",gap:"5px"}}>
          {fatias.map((f,i)=>(
            <div key={f.nome}
              style={{display:"flex",alignItems:"center",gap:"8px",padding:"5px 8px",borderRadius:T.radiusSm,
                background:hover===i?T.bg:"transparent",cursor:"pointer",transition:"background .12s"}}
              onMouseEnter={()=>setHover(i)} onMouseLeave={()=>setHover(null)}>
              <div style={{width:"10px",height:"10px",borderRadius:"3px",background:f.color,flexShrink:0}}/>
              <span style={{flex:1,fontSize:"12.5px",color:T.text,fontWeight:hover===i?600:400,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{f.nome}</span>
              <span style={{fontSize:"12px",fontFamily:T.mono,color:T.text,fontWeight:600,flexShrink:0}}>{fmt(f.valor)}</span>
              <span style={{fontSize:"11px",color:T.textMuted,width:"32px",textAlign:"right",flexShrink:0}}>{Math.round(f.pct*100)}%</span>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════
//  FILTRO ESTILO EXCEL
// ═══════════════════════════════════════════════════════════════════════════
// ═══════════════════════════════════════════════════════════════════════════
//  FILTRO DE DATA (intervalo)
// ═══════════════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════════════
function HomePage({ notas, tarefas, setTarefas, onVerDetalhes }) {
  const td=today(), tm=tomorrow(), ew=endOfWeek();

  // Calcula o valor das PARCELAS que vencem em cada período (não o total da nota)
  function somaParcelasNoPeriodo(dataFn) {
    let soma = 0;
    notas.forEach(n => {
      if (statusNota(n).key === "quitado") return;
      n.vencimentos.forEach(v => {
        if (parcelaStatusEfetivo(n,v) === "pago") return;
        if (dataFn(v)) soma += n.valor / n.parcelas;
      });
    });
    return soma;
  }
  function notasComParcelaNoperiodo(dataFn) {
    return notas.filter(n => {
      if (statusNota(n).key === "quitado") return false;
      return n.vencimentos.some(v => parcelaStatusEfetivo(n,v) !== "pago" && dataFn(v));
    });
  }

  function contarParcelasNoPeriodo(dataFn) {
    let count = 0;
    notas.forEach(n => {
      if (statusNota(n).key === "quitado") return;
      n.vencimentos.forEach(v => {
        if (parcelaStatusEfetivo(n,v) !== "pago" && dataFn(v)) count++;
      });
    });
    return count;
  }

  const parcelasHoje   = contarParcelasNoPeriodo(v => v === td);
  const parcelasAmanha = contarParcelasNoPeriodo(v => v === tm);
  const parcelasSemana = contarParcelasNoPeriodo(v => v >= td && v <= ew);

  function exportXLSX() {
    // Descobrir max parcelas
    const maxParc = Math.max(...filtered.map(n => n.parcelas || 1));

    // Cabeçalho dinâmico
    const hdr = [
      "FORNECEDOR","EMPRESA","NOTA FISCAL","DATA DE EMISSÃO",
      "VALOR DA NOTA","BOLETO RECEBIDO","PARCELAS",
      "VENCIMENTO","VALOR DA PARCELA",
    ];
    for (let i = 2; i <= maxParc; i++) hdr.push(`VENCIMENTO P${i}`);

    const rows = filtered.map(n => {
      const parcVal = n.parcelas > 0 ? (n.valor / n.parcelas) : n.valor;
      const row = [
        fnome(n.fornecedor),
        n.empresa,
        n.numero,
        fdate(n.emissao),
        n.valor,
        n.boletosRecebidos ? "BOLETO RECEBIDO" : "BOLETO NÃO RECEBIDO",
        n.parcelas,
        fdate(n.vencimentos[0]) || "",
        n.parcelas > 1 ? parcVal : "",
      ];
      for (let i = 1; i < maxParc; i++) {
        row.push(fdate(n.vencimentos[i]) || "");
      }
      return row;
    });

    // Montar worksheet
    const wsData = [hdr, ...rows];
    const XLSX = window.XLSX;
    const ws = XLSX.utils.aoa_to_sheet(wsData);

    // Larguras de coluna
    ws["!cols"] = [
      {wch:30},{wch:14},{wch:14},{wch:16},
      {wch:16},{wch:22},{wch:10},
      {wch:16},{wch:18},
      ...Array(Math.max(0, maxParc-1)).fill({wch:16}),
    ];

    // Estilo do cabeçalho (azul escuro + texto branco)
    const headerStyle = {
      font: { bold: true, color: { rgb: "FFFFFF" }, name: "Arial", sz: 11 },
      fill: { fgColor: { rgb: "1A5173" }, patternType: "solid" },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: {
        top:    { style: "thin", color: { rgb: "CCCCCC" } },
        bottom: { style: "thin", color: { rgb: "CCCCCC" } },
        left:   { style: "thin", color: { rgb: "CCCCCC" } },
        right:  { style: "thin", color: { rgb: "CCCCCC" } },
      }
    };

    // Estilo linha par (branco)
    const evenStyle = {
      font: { name: "Arial", sz: 10, color: { rgb: "111111" } },
      fill: { fgColor: { rgb: "FFFFFF" }, patternType: "solid" },
      alignment: { vertical: "center" },
      border: {
        top:    { style: "thin", color: { rgb: "E2E8F0" } },
        bottom: { style: "thin", color: { rgb: "E2E8F0" } },
        left:   { style: "thin", color: { rgb: "E2E8F0" } },
        right:  { style: "thin", color: { rgb: "E2E8F0" } },
      }
    };

    // Estilo linha ímpar (azul claro)
    const oddStyle = {
      ...evenStyle,
      fill: { fgColor: { rgb: "EEF4F9" }, patternType: "solid" },
    };

    // Estilo valor numérico
    const numStyle = (even) => ({
      ...(even ? evenStyle : oddStyle),
      alignment: { horizontal: "right", vertical: "center" },
      numFmt: "R$ #,##0.00",
    });

    // Aplicar estilos a todas as células
    const range = XLSX.utils.decode_range(ws["!ref"]);
    for (let R = range.s.r; R <= range.e.r; R++) {
      for (let C = range.s.c; C <= range.e.c; C++) {
        const addr = XLSX.utils.encode_cell({ r: R, c: C });
        if (!ws[addr]) ws[addr] = { v: "", t: "s" };
        if (R === 0) {
          ws[addr].s = headerStyle;
        } else {
          const isEven = R % 2 === 0;
          // Colunas de valor (4=valor nota, 8=valor parcela)
          if (C === 4 || C === 8) {
            ws[addr].s = numStyle(isEven);
            if (ws[addr].v !== "") { ws[addr].t = "n"; ws[addr].z = "R$ #,##0.00"; }
          } else {
            ws[addr].s = isEven ? evenStyle : oddStyle;
          }
        }
      }
    }

    // Congelar primeira linha
    ws["!freeze"] = { xSplit: 0, ySplit: 1 };

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Notas Fiscais");

    // Linha de totais
    const totalRow = Array(hdr.length).fill("");
    totalRow[0] = "TOTAL";
    totalRow[4] = filtered.reduce((s, n) => s + (n.valor || 0), 0);
    XLSX.utils.sheet_add_aoa(ws, [totalRow], { origin: -1 });
    const totalRowIdx = wsData.length;
    hdr.forEach((_, C) => {
      const addr = XLSX.utils.encode_cell({ r: totalRowIdx, c: C });
      if (!ws[addr]) ws[addr] = { v: "", t: "s" };
      ws[addr].s = {
        font: { bold: true, name: "Arial", sz: 11, color: { rgb: "1A5173" } },
        fill: { fgColor: { rgb: "C4DDF2" }, patternType: "solid" },
        alignment: C === 4 ? { horizontal: "right" } : {},
        border: headerStyle.border,
      };
      if (C === 4) { ws[addr].t = "n"; ws[addr].z = "R$ #,##0.00"; }
    });

    const hoje = new Date().toISOString().split("T")[0];
    XLSX.writeFile(wb, `notas_fiscais_${hoje}.xlsx`);
  }

  const vencHoje   = notasComParcelaNoperiodo(v => v === td);
  const vencAmanha = notasComParcelaNoperiodo(v => v === tm);
  const vencSemana = notasComParcelaNoperiodo(v => v >= td && v <= ew);
  const totalHoje   = somaParcelasNoPeriodo(v => v === td);
  const totalAmanha = somaParcelasNoPeriodo(v => v === tm);
  const totalSemana = somaParcelasNoPeriodo(v => v >= td && v <= ew);

  // Gráfico por empresa — período selecionável
  const COLORS = ["#F24E29","#5B89A6","#1A5173"];
  const [periodoEmpresa, setPeriodoEmpresa] = React.useState("mes"); // "hoje" | "semana" | "mes"
  const mesAtualKey  = `${new Date().getFullYear()}-${String(new Date().getMonth()+1).padStart(2,"0")}`;
  const mesAtualNome = new Date().toLocaleString("pt-BR",{month:"long"});

  const periodoFn = {
    hoje:  v => v === td,
    semana: v => v >= td && v <= ew,
    mes:   v => v.startsWith(mesAtualKey),
  }[periodoEmpresa];

  const periodoLabel = { hoje:"Hoje", semana:"Esta semana", mes: mesAtualNome }[periodoEmpresa];

  const porEmpresa = EMPRESAS.map((emp,i)=>{
    let total = 0, count = 0;
    notas.filter(n=>n.empresa===emp).forEach(n=>{
      n.vencimentos.forEach(v=>{
        if(periodoFn(v) && parcelaStatusEfetivo(n,v)!=="pago") {
          total += n.valor/n.parcelas;
          count++;
        }
      });
    });
    return { name:emp, color:COLORS[i], total, count };
  });
  const maxEmp = Math.max(...porEmpresa.map(e=>e.total),1);

  // (fluxo 3 meses removido — substituído por gráfico diário)

  return (
    <div style={{padding:"36px 40px",fontFamily:T.font,maxWidth:"1200px"}}>
      <div style={{marginBottom:"28px"}}>
        <h1 style={{margin:0,fontSize:"22px",fontWeight:700,color:T.text,letterSpacing:"-.02em"}}>Dashboard</h1>
        <p style={{margin:"4px 0 0",fontSize:"13.5px",color:T.textMuted}}>Controle de notas fiscais e boletos — Kowalsky</p>
      </div>

      {/* Cards de vencimento — valor das parcelas do período */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:"12px",marginBottom:"20px"}}>
        {[
          {label:"Hoje",   notas:vencHoje,   total:totalHoje,   parcelas:parcelasHoje,   accent:"#F24E29", accentBg:"#fdecea", tipo:"hoje"},
          {label:"Amanhã", notas:vencAmanha, total:totalAmanha, parcelas:parcelasAmanha, accent:T.primary, accentBg:T.primaryLight, tipo:"amanha"},
          {label:"Semana", notas:vencSemana, total:totalSemana, parcelas:parcelasSemana, accent:T.primary, accentBg:T.primaryLight, tipo:"semana"},
        ].map(c=>(
          <div key={c.label} onClick={()=>c.notas.length>0&&onVerDetalhes(c.tipo,c.label,c.notas)}
            style={{background:T.surface,borderRadius:T.radius,padding:"18px 20px",boxShadow:T.shadow,border:`1px solid ${T.border}`,cursor:c.notas.length>0?"pointer":"default",transition:"box-shadow .15s,transform .15s"}}
            onMouseEnter={e=>{if(c.notas.length>0){e.currentTarget.style.boxShadow=T.shadowMd;e.currentTarget.style.transform="translateY(-1px)";}}}
            onMouseLeave={e=>{e.currentTarget.style.boxShadow=T.shadow;e.currentTarget.style.transform="";}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:"12px"}}>
              <span style={{fontSize:"12px",fontWeight:500,color:T.textMuted}}>Vencimentos — {c.label}</span>
              {c.parcelas>0&&<span style={{fontSize:"11px",fontWeight:600,padding:"2px 7px",borderRadius:"20px",background:c.accentBg,color:c.accent}}>{c.parcelas} parcela(s)</span>}
            </div>
            <p style={{margin:"0 0 4px",fontSize:"22px",fontWeight:700,color:T.text,letterSpacing:"-.02em",fontFamily:T.mono}}>{fmt(c.total)}</p>
            {c.notas.length>0
              ? <p style={{margin:0,fontSize:"11.5px",color:T.textMuted}}>Ver detalhes →</p>
              : <p style={{margin:0,fontSize:"11.5px",color:T.textMuted}}>Nenhum vencimento ✓</p>}
          </div>
        ))}
      </div>

      {/* Linha inferior: Empresa (mês atual) + Gráfico diário */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 2fr",gap:"12px",marginBottom:"16px"}}>
        {/* Por empresa — período selecionável */}
        <div style={{background:T.surface,borderRadius:T.radius,padding:"20px 22px",boxShadow:T.shadow,border:`1px solid ${T.border}`}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:"16px"}}>
            <div>
              <p style={{margin:0,fontSize:"13px",fontWeight:600,color:T.text}}>Vencimentos por Empresa</p>
              <p style={{margin:"2px 0 0",fontSize:"12px",color:T.textMuted,textTransform:"capitalize"}}>{periodoLabel} — parcelas a pagar</p>
            </div>
            <div style={{display:"flex",gap:"3px"}}>
              {[{key:"hoje",label:"Hoje"},{key:"semana",label:"Semana"},{key:"mes",label:"Mês"}].map(p=>(
                <button key={p.key} onClick={()=>setPeriodoEmpresa(p.key)}
                  style={{padding:"3px 9px",borderRadius:T.radiusSm,border:`1.5px solid ${periodoEmpresa===p.key?T.primary:T.border}`,background:periodoEmpresa===p.key?T.primaryLight:T.surface,color:periodoEmpresa===p.key?T.primary:T.textMuted,fontSize:"11px",fontWeight:600,cursor:"pointer",fontFamily:T.font}}>
                  {p.label}
                </button>
              ))}
            </div>
          </div>
          <div style={{display:"flex",flexDirection:"column",gap:"14px"}}>
            {porEmpresa.map(e=>(
              <div key={e.name}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:"5px",alignItems:"center"}}>
                  <div style={{display:"flex",alignItems:"center",gap:"7px"}}>
                    <div style={{width:"7px",height:"7px",borderRadius:"50%",background:e.color}}/>
                    <span style={{fontSize:"12.5px",fontWeight:500,color:T.textSub}}>{e.name}</span>
                  </div>
                  <span style={{fontSize:"13px",fontWeight:600,color:T.text,fontFamily:T.mono}}>{fmt(e.total)}</span>
                </div>
                <div style={{height:"5px",borderRadius:"99px",background:T.bg,overflow:"hidden"}}>
                  <div style={{height:"100%",width:`${(e.total/maxEmp)*100}%`,borderRadius:"99px",background:e.color,transition:"width .6s"}}/>
                </div>
                <div style={{fontSize:"11px",color:T.textMuted,marginTop:"3px"}}>
                  {e.count > 0 ? `${e.count} parcela(s)` : "Nenhuma parcela"}
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* Gráfico diário de vencimentos */}
        <BarChartDiario notas={notas}/>
      </div>

      {/* Tarefas do dia */}
      {(() => {
        const tarefasHoje = tarefas.filter(t=>tarefaAtivaHoje(t));
        if (tarefasHoje.length===0) return null;
        return (
          <div style={{background:T.surface,borderRadius:T.radius,padding:"20px 22px",boxShadow:T.shadow,border:`1px solid ${T.border}`}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"14px"}}>
              <div>
                <p style={{margin:0,fontSize:"13px",fontWeight:600,color:T.text}}>Tarefas de Hoje</p>
                <p style={{margin:"2px 0 0",fontSize:"11.5px",color:T.textMuted}}>
                  {tarefasHoje.filter(t=>t.concluidoEm===today()).length}/{tarefasHoje.length} concluídas
                </p>
              </div>
              <div style={{height:"4px",width:"100px",borderRadius:"99px",background:T.bg,overflow:"hidden"}}>
                <div style={{height:"100%",borderRadius:"99px",background:"#17b26a",width:`${(tarefasHoje.filter(t=>t.concluidoEm===today()).length/tarefasHoje.length)*100}%`,transition:"width .4s"}}/>
              </div>
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:"6px"}}>
              {tarefasHoje.map(t=>{
                const concluida=t.concluidoEm===today();
                const rec=RECORRENCIA_CFG[t.recorrencia];
                return (
                  <div key={t.id} style={{display:"flex",alignItems:"center",gap:"10px",padding:"9px 12px",borderRadius:T.radiusSm,background:concluida?"#f6fef9":T.bg,border:`1px solid ${concluida?"#abefc6":T.border}`,transition:"all .15s"}}>
                    <button onClick={()=>setTarefas(ts=>ts.map(x=>x.id===t.id?{...x,concluidoEm:concluida?null:today()}:x))}
                      style={{width:"18px",height:"18px",borderRadius:"50%",border:`1.5px solid ${concluida?"#17b26a":T.borderStrong}`,background:concluida?"#17b26a":"transparent",cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,transition:"all .15s"}}>
                      {concluida&&<span style={{color:"#fff",fontSize:"10px",fontWeight:700}}>✓</span>}
                    </button>
                    <span style={{flex:1,fontSize:"13px",fontWeight:500,color:concluida?T.textMuted:T.text,textDecoration:concluida?"line-through":"none"}}>{t.titulo}</span>
                    <span style={{fontSize:"11px",fontWeight:500,padding:"2px 7px",borderRadius:"20px",background:rec.bg,color:rec.color}}>{rec.label}</span>
                  </div>
                );
              })}
            </div>
          </div>
        );
      })()}

      {/* Gráfico de Pizza */}
      <PieChartFornecedores notas={notas}/>
    </div>
  );
}
// ═══════════════════════════════════════════════════════════════════════════
//  FILTRO ESTILO EXCEL
// ═══════════════════════════════════════════════════════════════════════════
// ═══════════════════════════════════════════════════════════════════════════
//  FILTRO DE DATA (intervalo)
// ═══════════════════════════════════════════════════════════════════════════
function DateRangeFilter({ de, ate, setDe, setAte }) {
  const [open, setOpen] = useState(false);
  const [tmpDe,  setTmpDe]  = useState(de);
  const [tmpAte, setTmpAte] = useState(ate);

  // Sincroniza estado interno sempre que o dropdown abre ou o filtro externo muda
  React.useEffect(() => { if (open) { setTmpDe(de); setTmpAte(ate); } }, [open, de, ate]);

  function apply() { setDe(tmpDe); setAte(tmpAte); setOpen(false); }
  function clear()  { setTmpDe(""); setTmpAte(""); setDe(""); setAte(""); setOpen(false); }

  const isFiltered = de || ate;
  const label = isFiltered
    ? (de && ate ? `${fdate(de)} até ${fdate(ate)}` : de ? `A partir de ${fdate(de)}` : `Até ${fdate(ate)}`)
    : null;

  return (
    <div style={{position:"relative",display:"inline-block"}}>
      <button onClick={()=>{ setTmpDe(de); setTmpAte(ate); setOpen(o=>!o); }}
        title={isFiltered ? label : "Filtrar por data"}
        style={{background:"none",border:"none",cursor:"pointer",padding:"2px 4px",borderRadius:"4px",color:isFiltered?"#1A5173":"#94a3b8",fontSize:"12px"}}>
        {isFiltered ? "\uD83D\uDD35" : "\u22df"}
      </button>
      {open && (
        <>
          <div style={{position:"fixed",inset:0,zIndex:80}} onClick={()=>setOpen(false)}/>
          <div style={{position:"absolute",top:"calc(100% + 4px)",left:0,background:T.surface,borderRadius:"10px",boxShadow:"0 8px 28px rgba(0,0,0,0.14)",border:`1px solid ${T.border}`,zIndex:90,minWidth:"230px",padding:"12px"}}>
            <div style={{fontSize:"11px",fontWeight:700,color:T.textSub,textTransform:"uppercase",letterSpacing:".05em",marginBottom:"10px",paddingBottom:"6px",borderBottom:"1px solid #f1f5f9"}}>
              Filtrar: Data de Emissão
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:"8px"}}>
              <div>
                <label style={{fontSize:"11.5px",fontWeight:600,color:T.textSub,display:"block",marginBottom:"4px"}}>De:</label>
                <input type="date" value={tmpDe} onChange={e=>setTmpDe(e.target.value)}
                  style={{...S.input,fontSize:"13px",padding:"7px 10px",width:"100%",boxSizing:"border-box"}}/>
              </div>
              <div>
                <label style={{fontSize:"11.5px",fontWeight:600,color:T.textSub,display:"block",marginBottom:"4px"}}>Até:</label>
                <input type="date" value={tmpAte} onChange={e=>setTmpAte(e.target.value)}
                  style={{...S.input,fontSize:"13px",padding:"7px 10px",width:"100%",boxSizing:"border-box"}}/>
              </div>
            </div>
            <div style={{display:"flex",gap:"5px",marginTop:"10px",flexWrap:"wrap"}}>
              {[
                {label:"Este mes", onClick:()=>{ const n=new Date(); const y=n.getFullYear(); const m=String(n.getMonth()+1).padStart(2,"0"); setTmpDe(`${y}-${m}-01`); setTmpAte(`${y}-${m}-31`); }},
                {label:"Este ano", onClick:()=>{ const y=new Date().getFullYear(); setTmpDe(`${y}-01-01`); setTmpAte(`${y}-12-31`); }},
                {label:"Hoje",     onClick:()=>{ const d=today(); setTmpDe(d); setTmpAte(d); }},
              ].map(a=>(
                <button key={a.label} onClick={a.onClick} style={{padding:"4px 9px",borderRadius:"6px",border:"1.5px solid #e2e8f0",background:T.bg,color:"#475569",fontSize:"11.5px",fontWeight:600,cursor:"pointer"}}>{a.label}</button>
              ))}
            </div>
            <div style={{display:"flex",gap:"6px",marginTop:"10px",paddingTop:"8px",borderTop:"1px solid #f1f5f9"}}>
              <button onClick={clear} style={{flex:1,padding:"6px",borderRadius:"7px",border:"1.5px solid #e2e8f0",background:T.bg,color:"#475569",fontWeight:600,fontSize:"12px",cursor:"pointer"}}>Limpar</button>
              <button onClick={apply} style={{flex:2,padding:"6px",borderRadius:"7px",border:"none",background:"#1A5173",color:"#fff",fontWeight:700,fontSize:"12px",cursor:"pointer"}}>Aplicar</button>
            </div>
          </div>
        </>
      )}
    </div>
  );
}

function ColFilter({ label, values, active, onApply, onClear }) {
  const [open, setOpen]     = useState(false);
  const [search, setSearch] = useState("");
  const [sel, setSel]       = useState(active || []);

  React.useEffect(() => { if (open) setSel(active || []); }, [open]);

  const filteredVals = values.filter(v => v.toLowerCase().includes(search.toLowerCase()));
  const allSel = sel.length === 0 || sel.length === values.length;

  function toggle(v) { setSel(s => s.includes(v) ? s.filter(x=>x!==v) : [...s,v]); }
  function toggleAll() { setSel(allSel ? [] : [...values]); }
  function apply() { onApply(sel.length===values.length||sel.length===0 ? [] : sel); setOpen(false); }
  function clear() { setSel([]); onClear(); setOpen(false); }

  const isFiltered = active && active.length > 0;

  return (
    <div style={{position:"relative",display:"inline-block"}}>
      <button onClick={()=>{ setSearch(""); setOpen(o=>!o); }}
        title={isFiltered?"Filtro ativo":"Filtrar"}
        style={{background:"none",border:"none",cursor:"pointer",padding:"2px 4px",borderRadius:"4px",color:isFiltered?"#1A5173":"#94a3b8",fontSize:"12px"}}>
        {isFiltered ? "🔵" : "⊟"}
      </button>
      {open && (
        <>
          <div style={{position:"fixed",inset:0,zIndex:80}} onClick={()=>setOpen(false)}/>
          <div style={{position:"absolute",top:"calc(100% + 4px)",left:0,background:T.surface,borderRadius:"10px",boxShadow:"0 8px 28px rgba(0,0,0,0.14)",border:`1px solid ${T.border}`,zIndex:90,minWidth:"220px",padding:"10px"}}>
            <div style={{fontSize:"11px",fontWeight:700,color:T.textSub,textTransform:"uppercase",letterSpacing:".05em",marginBottom:"8px",paddingBottom:"6px",borderBottom:"1px solid #f1f5f9"}}>
              Filtrar: {label}
            </div>
            <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Pesquisar..."
              style={{...S.input,marginBottom:"8px",fontSize:"12px",padding:"6px 10px"}} autoFocus/>
            <div style={{maxHeight:"180px",overflowY:"auto",display:"flex",flexDirection:"column",gap:"2px"}}>
              <label style={{display:"flex",alignItems:"center",gap:"8px",padding:"5px 6px",borderRadius:"6px",cursor:"pointer",fontSize:"13px",fontWeight:600,color:T.text}}
                onMouseEnter={e=>e.currentTarget.style.background="#f8fafc"}
                onMouseLeave={e=>e.currentTarget.style.background=""}>
                <input type="checkbox" checked={allSel} onChange={toggleAll} style={{accentColor:"#1A5173",width:"14px",height:"14px"}}/>
                (Selecionar Tudo)
              </label>
              {filteredVals.map(v=>(
                <label key={v} style={{display:"flex",alignItems:"center",gap:"8px",padding:"5px 6px",borderRadius:"6px",cursor:"pointer",fontSize:"13px",color:"#374151"}}
                  onMouseEnter={e=>e.currentTarget.style.background="#f8fafc"}
                  onMouseLeave={e=>e.currentTarget.style.background=""}>
                  <input type="checkbox" checked={allSel||sel.includes(v)} onChange={()=>toggle(v)} style={{accentColor:"#1A5173",width:"14px",height:"14px"}}/>
                  {v}
                </label>
              ))}
              {filteredVals.length===0 && <div style={{padding:"8px 6px",fontSize:"12px",color:T.textMuted}}>Sem resultados</div>}
            </div>
            <div style={{display:"flex",gap:"6px",marginTop:"10px",paddingTop:"8px",borderTop:"1px solid #f1f5f9"}}>
              <button onClick={clear} style={{flex:1,padding:"6px",borderRadius:"7px",border:"1.5px solid #e2e8f0",background:T.bg,color:"#475569",fontWeight:600,fontSize:"12px",cursor:"pointer"}}>Limpar</button>
              <button onClick={apply} style={{flex:2,padding:"6px",borderRadius:"7px",border:"none",background:"#1A5173",color:"#fff",fontWeight:700,fontSize:"12px",cursor:"pointer"}}>Aplicar</button>
            </div>
          </div>
        </>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════
function NotasFiscaisPage({ notas, setNotas, showModal, setShowModal, fornecedores, onNovoFornecedor, lastAddedId, setLastAddedId }) {
  const { toasts, toast } = useToast();
  const [editNota,setEditNota]     = useState(null);
  const [deleteId,setDeleteId]     = useState(null);
  const [expandedId,setExpandedId] = useState(null);
  const [search,setSearch]         = useState("");
  const [page, setPage]            = useState(1);
  const PER_PAGE = 50;
  const toggleExpand = (id) => setExpandedId(prev => prev===id ? null : id);

  // Filtros por coluna
  const [fNumero,  setFNumero]  = useState([]);
  const [fFornec,  setFFornec]  = useState([]);
  const [fEmpresa, setFEmpresa] = useState([]);
  const [fStatus,  setFStatus]  = useState([]);
  const [fValor,   setFValor]   = useState([]);
  const [fDataDe,  setFDataDe]  = useState("");
  const [fDataAte, setFDataAte] = useState("");
  const [sortCol,  setSortCol]  = useState(null);   // "numero"|"fornecedor"|"emissao"|"empresa"|"valor"|"parcelas"|"status"
  const [sortDir,  setSortDir]  = useState("asc");  // "asc"|"desc"

  function toggleSort(col) {
    setSortCol(prev => {
      if (prev === col) { setSortDir(d => d === "asc" ? "desc" : "asc"); return col; }
      setSortDir("asc"); return col;
    });
    setPage(1);
  }

  const uNumero  = useMemo(()=>[...new Set(notas.map(n=>n.numero))].sort(), [notas]);
  const uFornec  = useMemo(()=>[...new Set(notas.map(n=>n.fornecedor))].sort(), [notas]);
  const uEmpresa = useMemo(()=>[...new Set(notas.map(n=>n.empresa))].sort(), [notas]);
  const uStatus  = useMemo(()=>[...new Set(notas.map(n=>statusNota(n).label))].sort(), [notas]);
  const uValor   = useMemo(()=>[...new Set(notas.map(n=>fmt(n.valor)))].sort((a,b)=>{
    const pa=parseFloat(a.replace(/[R$\s.]/g,"").replace(",","."));
    const pb=parseFloat(b.replace(/[R$\s.]/g,"").replace(",","."));
    return pa-pb;
  }), [notas]);

  const filtered = useMemo(()=>notas.filter(n=>{
    if(fNumero.length  && !fNumero.includes(n.numero))                    return false;
    if(fFornec.length  && !fFornec.includes(n.fornecedor))                return false;
    if(fEmpresa.length && !fEmpresa.includes(n.empresa))                  return false;
    if(fStatus.length  && !fStatus.includes(statusNota(n).label)) return false;
    if(fValor.length   && !fValor.includes(fmt(n.valor)))                 return false;
    if(fDataDe  && n.emissao < fDataDe)  return false;
    if(fDataAte && n.emissao > fDataAte) return false;
    if(search && ![n.fornecedor,n.numero,n.empresa].some(x=>x.toLowerCase().includes(search.toLowerCase()))) return false;
    return true;
  }),[notas,fNumero,fFornec,fEmpresa,fStatus,fValor,fDataDe,fDataAte,search]);

  const hasFilter = fNumero.length||fFornec.length||fEmpresa.length||fStatus.length||fValor.length||fDataDe||fDataAte||search;
  function clearAll() { setFNumero([]); setFFornec([]); setFEmpresa([]); setFStatus([]); setFValor([]); setFDataDe(""); setFDataAte(""); setSearch(""); setPage(1); }

  // Reset page when filters change
  React.useEffect(() => setPage(1), [fNumero,fFornec,fEmpresa,fStatus,fValor,fDataDe,fDataAte,search]);

  const sorted = useMemo(() => {
    if (!sortCol) return filtered;
    return [...filtered].sort((a, b) => {
      let va, vb;
      if (sortCol === "numero")     { va = a.numero;                      vb = b.numero; }
      else if (sortCol === "fornecedor") { va = a.fornecedor;             vb = b.fornecedor; }
      else if (sortCol === "emissao")   { va = a.emissao;                 vb = b.emissao; }
      else if (sortCol === "empresa")   { va = a.empresa;                 vb = b.empresa; }
      else if (sortCol === "valor")     { va = a.valor;                   vb = b.valor; return sortDir==="asc" ? va-vb : vb-va; }
      else if (sortCol === "parcelas")  { va = a.parcelas;                vb = b.parcelas; return sortDir==="asc" ? va-vb : vb-va; }
      else if (sortCol === "status")    { va = statusNota(a).label; vb = statusNota(b).label; }
      else return 0;
      return sortDir === "asc" ? va.localeCompare(vb, "pt-BR") : vb.localeCompare(va, "pt-BR");
    });
  }, [filtered, sortCol, sortDir]);

  const totalPages  = Math.max(1, Math.ceil(sorted.length / PER_PAGE));
  const paginated   = sorted.slice((page-1)*PER_PAGE, page*PER_PAGE);

  function handleSave(nota) {
    const isEdit = nota.id && notas.find(n=>n.id===nota.id);
    setNotas(ns=>isEdit?ns.map(n=>n.id===nota.id?nota:n):[...ns,nota]);
    setShowModal(false); setEditNota(null);
    toast(isEdit?"Nota fiscal atualizada":"Nota fiscal criada com sucesso");
    if (!isEdit && setLastAddedId) { setLastAddedId(nota.id); setTimeout(()=>setLastAddedId(null), 3000); }
  }
  function exportCSV() {
    if (!window.XLSX) { alert("Aguarde o carregamento do Excel..."); return; }
    const maxParc = Math.max(...filtered.map(n => n.parcelas || 1));
    const hdr = [
      "FORNECEDOR","EMPRESA","NOTA FISCAL","DATA DE EMISSÃO",
      "VALOR DA NOTA","BOLETO RECEBIDO","PARCELAS",
      "VENCIMENTO","VALOR DA PARCELA",
    ];
    for (let i = 2; i <= maxParc; i++) hdr.push(`VENCIMENTO P${i}`);
    const rows = filtered.map(n => {
      const parcVal = n.parcelas > 1 ? (n.valor / n.parcelas) : "";
      const row = [
        fnome(n.fornecedor), n.empresa, n.numero, fdate(n.emissao),
        n.valor, n.boletosRecebidos ? "BOLETO RECEBIDO" : "BOLETO NÃO RECEBIDO",
        n.parcelas, fdate(n.vencimentos[0]) || "", parcVal,
      ];
      for (let i = 1; i < maxParc; i++) row.push(fdate(n.vencimentos[i]) || "");
      return row;
    });
    const wsData = [hdr, ...rows];
    const XLSX = window.XLSX;
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    ws["!cols"] = [{wch:30},{wch:14},{wch:14},{wch:16},{wch:16},{wch:22},{wch:10},{wch:16},{wch:18},...Array(Math.max(0,maxParc-1)).fill({wch:16})];
    const headerStyle = {
      font:{bold:true,color:{rgb:"FFFFFF"},name:"Arial",sz:11},
      fill:{fgColor:{rgb:"1A5173"},patternType:"solid"},
      alignment:{horizontal:"center",vertical:"center",wrapText:true},
      border:{top:{style:"thin",color:{rgb:"CCCCCC"}},bottom:{style:"thin",color:{rgb:"CCCCCC"}},left:{style:"thin",color:{rgb:"CCCCCC"}},right:{style:"thin",color:{rgb:"CCCCCC"}}}
    };
    const evenStyle = {font:{name:"Arial",sz:10,color:{rgb:"111111"}},fill:{fgColor:{rgb:"FFFFFF"},patternType:"solid"},alignment:{vertical:"center"},border:{top:{style:"thin",color:{rgb:"E2E8F0"}},bottom:{style:"thin",color:{rgb:"E2E8F0"}},left:{style:"thin",color:{rgb:"E2E8F0"}},right:{style:"thin",color:{rgb:"E2E8F0"}}}};
    const oddStyle = {...evenStyle,fill:{fgColor:{rgb:"EEF4F9"},patternType:"solid"}};
    const range = XLSX.utils.decode_range(ws["!ref"]);
    for (let R=range.s.r;R<=range.e.r;R++) {
      for (let C=range.s.c;C<=range.e.c;C++) {
        const addr=XLSX.utils.encode_cell({r:R,c:C});
        if(!ws[addr]) ws[addr]={v:"",t:"s"};
        if(R===0){ ws[addr].s=headerStyle; }
        else {
          const isEven=R%2===0;
          if(C===4||C===8){ ws[addr].s={...(isEven?evenStyle:oddStyle),alignment:{horizontal:"right",vertical:"center"}}; if(ws[addr].v!==""){ws[addr].t="n";ws[addr].z="R$ #,##0.00";} }
          else { ws[addr].s=isEven?evenStyle:oddStyle; }
        }
      }
    }
    ws["!freeze"]={xSplit:0,ySplit:1};
    // Linha de totais
    const totalRow=Array(hdr.length).fill("");
    totalRow[0]="TOTAL"; totalRow[4]=filtered.reduce((s,n)=>s+(n.valor||0),0);
    XLSX.utils.sheet_add_aoa(ws,[totalRow],{origin:-1});
    const totalRowIdx=wsData.length;
    hdr.forEach((_,C)=>{
      const addr=XLSX.utils.encode_cell({r:totalRowIdx,c:C});
      if(!ws[addr]) ws[addr]={v:"",t:"s"};
      ws[addr].s={font:{bold:true,name:"Arial",sz:11,color:{rgb:"1A5173"}},fill:{fgColor:{rgb:"C4DDF2"},patternType:"solid"},alignment:C===4?{horizontal:"right"}:{},border:headerStyle.border};
      if(C===4){ws[addr].t="n";ws[addr].z="R$ #,##0.00";}
    });
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Notas Fiscais");
    const hoje=new Date().toISOString().split("T")[0];
    XLSX.writeFile(wb,`notas_fiscais_${hoje}.xlsx`);
  }

  const vencHoje = notas.filter(n=>statusNota(n).key!=="quitado"&&n.vencimentos.some(v=>v===today())).length;

  const COLS = [
    {label:"Número NF",   sortKey:"numero",     active:fNumero,  onApply:setFNumero,  onClear:()=>setFNumero([]),  vals:uNumero },
    {label:"Fornecedor",  sortKey:"fornecedor",  active:fFornec,  onApply:setFFornec,  onClear:()=>setFFornec([]),  vals:uFornec },
    {label:"Data Emissão",sortKey:"emissao",     active:null,     onApply:null,        onClear:null,                vals:null,   dateRange:{de:fDataDe,ate:fDataAte,setDe:setFDataDe,setAte:setFDataAte} },
    {label:"Empresa",     sortKey:"empresa",     active:fEmpresa, onApply:setFEmpresa, onClear:()=>setFEmpresa([]), vals:uEmpresa},
    {label:"Valor Total", sortKey:"valor",       active:fValor,   onApply:setFValor,   onClear:()=>setFValor([]),   vals:uValor  },
    {label:"Parcelas",    sortKey:"parcelas",    active:null,     onApply:null,        onClear:null,                vals:null    },
    {label:"Status",      sortKey:"status",      active:fStatus,  onApply:setFStatus,  onClear:()=>setFStatus([]),  vals:uStatus },
    {label:"Ações",       sortKey:null,          active:null,     onApply:null,        onClear:null,                vals:null    },
  ];

  return (
    <div style={{padding:"36px 40px",fontFamily:T.font,maxWidth:"1400px"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:"24px"}}>
        <div>
          <h1 style={{margin:0,fontSize:"22px",fontWeight:700,color:T.text,letterSpacing:"-.02em"}}>Notas Fiscais</h1>
          <p style={{margin:"4px 0 0",fontSize:"13px",color:T.textMuted}}>
            {filtered.length} de {notas.length} nota(s)
            {hasFilter ? <button onClick={clearAll} style={{marginLeft:"10px",fontSize:"12px",color:"#F24E29",background:"none",border:"none",cursor:"pointer",fontWeight:600}}>✕ Limpar filtros</button> : ""}
          </p>
        </div>
        <div style={{display:"flex",gap:"8px"}}>
          <button onClick={exportCSV} style={{padding:"8px 14px",borderRadius:T.radiusSm,border:`1.5px solid ${T.border}`,background:T.surface,color:T.textSub,fontWeight:500,fontSize:"13px",cursor:"pointer",fontFamily:T.font}}>↓ CSV</button>
          <button onClick={()=>setShowModal(true)} style={{padding:"8px 16px",borderRadius:T.radiusSm,border:"none",background:T.primary,color:"#fff",fontWeight:600,fontSize:"13px",cursor:"pointer",fontFamily:T.font}}>+ Nova NF</button>
        </div>
      </div>

      {vencHoje>0&&(
        <div style={{background:"#fffaeb",border:`1px solid #fec84b`,borderRadius:T.radiusSm,padding:"10px 14px",marginBottom:"16px",fontSize:"13px",color:"#92400e",display:"flex",alignItems:"center",gap:"8px"}}>
          <span style={{fontWeight:600}}>⚠ {vencHoje} nota(s) com vencimento hoje</span>
        </div>
      )}

      <div style={{marginBottom:"12px"}}>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar por NF, fornecedor ou empresa..." style={{...S.input,maxWidth:"360px",background:T.surface,boxShadow:T.shadow}}/>
      </div>

      <div style={{background:T.surface,borderRadius:T.radius,boxShadow:T.shadow,border:`1px solid ${T.border}`,overflow:"hidden"}}>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:"13px",fontFamily:T.font}}>
            <thead>
              <tr style={{background:T.bg}}>
                <th style={{padding:"10px 12px",width:"32px",borderBottom:`1px solid ${T.border}`}}></th>
                {COLS.map(col=>{
                  const isActive = sortCol === col.sortKey && col.sortKey;
                  const arrow = isActive ? (sortDir === "asc" ? " ▲" : " ▼") : "";
                  return (
                    <th key={col.label} style={{padding:"10px 14px",textAlign:"left",fontWeight:600,color:isActive?T.primary:T.textMuted,fontSize:"11px",textTransform:"uppercase",letterSpacing:".06em",borderBottom:`1px solid ${T.border}`,whiteSpace:"nowrap",userSelect:"none"}}>
                      <div style={{display:"flex",alignItems:"center",gap:"4px"}}>
                        {col.sortKey ? (
                          <span onClick={()=>toggleSort(col.sortKey)} style={{cursor:"pointer",display:"flex",alignItems:"center",gap:"2px"}}
                            title={`Ordenar por ${col.label}`}>
                            {col.label}
                            <span style={{fontSize:"10px",color:isActive?T.primary:T.border,marginLeft:"1px"}}>
                              {isActive ? arrow : " ⇅"}
                            </span>
                          </span>
                        ) : col.label}
                        {col.vals && <ColFilter label={col.label} values={col.vals} active={col.active} onApply={col.onApply} onClear={col.onClear}/>}
                        {col.dateRange && <DateRangeFilter de={col.dateRange.de} ate={col.dateRange.ate} setDe={col.dateRange.setDe} setAte={col.dateRange.setAte}/>}
                      </div>
                    </th>
                  );
                })}
              </tr>
            </thead>
            <tbody>
              {filtered.length===0?(
                <tr><td colSpan={9} style={{padding:"48px",textAlign:"center",color:T.textMuted,fontSize:"13px"}}>Nenhuma nota fiscal encontrada</td></tr>
              ):paginated.map((n,i)=>{
                const st=statusNota(n);
                const expanded=expandedId===n.id;
                const isLast=i===paginated.length-1;
                return (
                  <React.Fragment key={n.id}>
                    <tr onClick={()=>toggleExpand(n.id)}
                      style={{borderBottom:expanded?"none":isLast?"none":`1px solid ${T.border}`,cursor:"pointer",transition:"background .1s"}}
                      onMouseEnter={e=>{ if(!expanded) e.currentTarget.style.background=T.bg; }}
                      onMouseLeave={e=>{ if(!expanded) e.currentTarget.style.background=""; }}>
                      <td style={{padding:"12px 12px",textAlign:"center"}}>
                        <span style={{fontSize:"10px",color:T.textMuted,display:"inline-block",transform:expanded?"rotate(90deg)":"rotate(0deg)",transition:"transform .18s"}}>▶</span>
                      </td>
                      <td style={{padding:"12px 14px",fontWeight:600,color:T.text,fontFamily:T.mono,fontSize:"12.5px"}}>{n.numero}</td>
                      <td style={{padding:"12px 14px",fontWeight:500,color:T.text}}>{fnome(n.fornecedor)}</td>
                      <td style={{padding:"12px 14px",color:T.textSub}}>{fdate(n.emissao)}</td>
                      <td style={{padding:"12px 14px"}}><span style={{background:T.bg,color:T.textSub,borderRadius:"5px",padding:"2px 8px",fontSize:"11.5px",fontWeight:500,border:`1px solid ${T.border}`}}>{n.empresa}</span></td>
                      <td style={{padding:"12px 14px",fontWeight:600,color:T.text,fontFamily:T.mono,fontSize:"12.5px"}}>{fmt(n.valor)}</td>
                      <td style={{padding:"12px 14px",color:T.textSub,textAlign:"center"}}>{n.parcelas}</td>
                      <td style={{padding:"12px 14px"}}><span style={{fontSize:"11.5px",fontWeight:500,padding:"2px 9px",borderRadius:"20px",background:st.bg,color:st.color}}>{st.label}</span></td>
                      <td style={{padding:"12px 14px"}} onClick={e=>e.stopPropagation()}>
                        <div style={{display:"flex",gap:"4px",alignItems:"center"}}>
                          {n.observacao && (
                            <div style={{position:"relative",display:"inline-flex"}}
                              onMouseEnter={e=>{ const t=e.currentTarget.querySelector(".obs-tt"); if(t) t.style.opacity="1"; }}
                              onMouseLeave={e=>{ const t=e.currentTarget.querySelector(".obs-tt"); if(t) t.style.opacity="0"; }}>
                              <button title={n.observacao}
                                style={{padding:"4px 8px",borderRadius:"6px",border:`1.5px solid ${T.border}`,background:"#f0f6ff",cursor:"pointer",fontSize:"13px",lineHeight:1,color:"#1A5173"}}>💬</button>
                              <div className="obs-tt" style={{position:"absolute",bottom:"calc(100% + 6px)",left:"50%",transform:"translateX(-50%)",background:"#1e293b",color:"#f8fafc",fontSize:"12px",padding:"7px 11px",borderRadius:"8px",whiteSpace:"pre-wrap",maxWidth:"220px",lineHeight:"1.4",zIndex:99,boxShadow:"0 4px 14px rgba(0,0,0,.25)",pointerEvents:"none",opacity:0,transition:"opacity .15s"}}>
                                {n.observacao}
                              </div>
                            </div>
                          )}
                          <button onClick={()=>setEditNota({...n,valor:n.valor.toLocaleString("pt-BR",{minimumFractionDigits:2,maximumFractionDigits:2})})} style={{padding:"4px 10px",borderRadius:"6px",border:`1.5px solid ${T.border}`,background:T.surface,cursor:"pointer",fontSize:"12px",fontWeight:500,color:T.textSub,fontFamily:T.font}}>Editar</button>
                          <button onClick={()=>setDeleteId(n.id)} style={{padding:"4px 10px",borderRadius:"6px",border:"1.5px solid #fdecea",background:"#fff5f5",cursor:"pointer",fontSize:"12px",fontWeight:500,color:"#F24E29",fontFamily:T.font}}>✕</button>
                        </div>
                      </td>
                    </tr>
                    {expanded && n.vencimentos.map((venc,pi)=>{
                      const isLastParcela=pi===n.vencimentos.length-1;
                      const valParcela=n.valor/n.parcelas;
                      const stEfetivo=parcelaStatusEfetivo(n,venc);
                      const isPago=stEfetivo==="pago";
                      const isManual=n.parcelasPagas&&n.parcelasPagas.includes(venc);
                      const isAuto=isPago&&!isManual&&n.boletosRecebidos&&venc<today();
                      const boletoCfg=n.boletosRecebidos?{label:"Boleto Recebido",color:"#067647",bg:"#ecfdf3"}:{label:"Boleto Pendente",color:"#b54708",bg:"#fffaeb"};

                      function toggleParcela(e) {
                        e.stopPropagation();
                        let wasPago = false;
                        setNotas(ns=>ns.map(x=>{
                          if(x.id!==n.id) return x;
                          const pagas = x.parcelasPagas ? [...x.parcelasPagas] : [];
                          const idx = pagas.indexOf(venc);
                          if(idx>=0) { pagas.splice(idx,1); wasPago=true; } else pagas.push(venc);
                          return {...x, parcelasPagas:pagas};
                        }));
                        toast(wasPago?`Parcela ${pi+1} desmarcada`:`Parcela ${pi+1} marcada como paga`);
                      }

                      return (
                        <tr key={`${n.id}-p${pi}`} style={{background:isPago?"#f6fef9":"#fafbff",borderBottom:isLastParcela&&isLast?"none":isLastParcela?`1px solid ${T.border}`:`1px solid ${T.bg}`}}>
                          <td style={{padding:"9px 12px"}}></td>
                          <td style={{padding:"9px 14px"}}><span style={{fontFamily:T.mono,fontSize:"12px",color:T.textMuted}}>{n.numero}-{pi+1}</span></td>
                          <td style={{padding:"9px 14px",fontSize:"12.5px",color:T.textSub}}>{fdate(venc)}</td>
                          <td colSpan={2} style={{padding:"9px 14px"}}>
                            <div style={{display:"flex",gap:"5px",alignItems:"center"}}>
                              <span style={{fontSize:"11.5px",fontWeight:500,padding:"2px 9px",borderRadius:"20px",background:boletoCfg.bg,color:boletoCfg.color}}>{boletoCfg.label}</span>
                              {isPago&&<span style={{fontSize:"11.5px",fontWeight:500,padding:"2px 9px",borderRadius:"20px",background:"#C4DDF2",color:T.primary}}>✓ Pago{isAuto?" (auto)":""}</span>}
                            </div>
                          </td>
                          <td style={{padding:"9px 14px",fontWeight:600,color:T.text,fontFamily:T.mono,fontSize:"12.5px"}}>{fmt(valParcela)}</td>
                          <td style={{padding:"9px 14px",fontSize:"12px",color:isPago?"#067647":venc===today()?"#c4780a":T.textMuted,fontWeight:isPago||venc===today()?600:400}}>
                            {isPago?(isAuto?"Pago automaticamente":"Pago manualmente"):venc===today()?"Vence hoje":venc<today()?"Vencido":"Vence "+fdate(venc)}
                          </td>
                          <td colSpan={2} style={{padding:"9px 14px",textAlign:"right"}} onClick={e=>e.stopPropagation()}>
                            {!isAuto&&(
                              <button onClick={toggleParcela}
                                style={{padding:"4px 12px",borderRadius:"6px",fontSize:"12px",fontWeight:500,cursor:"pointer",border:`1.5px solid ${isPago?"#abefc6":T.border}`,background:isPago?"#f6fef9":T.surface,color:isPago?"#067647":T.textSub,fontFamily:T.font,transition:"all .12s"}}>
                                {isPago?"↩ Desfazer":"✓ Marcar pago"}
                              </button>
                            )}
                          </td>
                        </tr>
                      );
                    })}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      {/* PAGINAÇÃO */}
      {totalPages > 1 && (
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginTop:"14px",padding:"0 2px"}}>
          <span style={{fontSize:"12.5px",color:T.textMuted,fontFamily:T.font}}>
            {(page-1)*PER_PAGE+1}–{Math.min(page*PER_PAGE,filtered.length)} de {filtered.length} notas
          </span>
          <div style={{display:"flex",gap:"3px",alignItems:"center"}}>
            <button onClick={()=>setPage(1)} disabled={page===1}
              style={{padding:"5px 9px",borderRadius:T.radiusSm,border:`1.5px solid ${T.border}`,background:T.surface,color:page===1?T.textMuted:T.textSub,fontWeight:600,fontSize:"12px",cursor:page===1?"default":"pointer",fontFamily:T.font}}>«</button>
            <button onClick={()=>setPage(p=>Math.max(1,p-1))} disabled={page===1}
              style={{padding:"5px 11px",borderRadius:T.radiusSm,border:`1.5px solid ${T.border}`,background:T.surface,color:page===1?T.textMuted:T.textSub,fontWeight:500,fontSize:"12.5px",cursor:page===1?"default":"pointer",fontFamily:T.font}}>‹</button>
            {Array.from({length:Math.min(5,totalPages)},(_,i)=>{
              let p;
              if(totalPages<=5) p=i+1;
              else if(page<=3) p=i+1;
              else if(page>=totalPages-2) p=totalPages-4+i;
              else p=page-2+i;
              return (
                <button key={p} onClick={()=>setPage(p)}
                  style={{padding:"5px 10px",borderRadius:T.radiusSm,border:page===p?"none":`1.5px solid ${T.border}`,background:page===p?T.primary:T.surface,color:page===p?"#fff":T.textSub,fontWeight:page===p?600:400,fontSize:"12.5px",cursor:"pointer",minWidth:"30px",fontFamily:T.font}}>
                  {p}
                </button>
              );
            })}
            <button onClick={()=>setPage(p=>Math.min(totalPages,p+1))} disabled={page===totalPages}
              style={{padding:"5px 11px",borderRadius:T.radiusSm,border:`1.5px solid ${T.border}`,background:T.surface,color:page===totalPages?T.textMuted:T.textSub,fontWeight:500,fontSize:"12.5px",cursor:page===totalPages?"default":"pointer",fontFamily:T.font}}>›</button>
            <button onClick={()=>setPage(totalPages)} disabled={page===totalPages}
              style={{padding:"5px 9px",borderRadius:T.radiusSm,border:`1.5px solid ${T.border}`,background:T.surface,color:page===totalPages?T.textMuted:T.textSub,fontWeight:600,fontSize:"12px",cursor:page===totalPages?"default":"pointer",fontFamily:T.font}}>»</button>
          </div>
        </div>
      )}

      {(showModal||editNota)&&<NFModal onClose={()=>{setShowModal(false);setEditNota(null);}} onSave={handleSave} editData={editNota} fornecedores={fornecedores} onNovoFornecedor={onNovoFornecedor}/>}
      {deleteId&&(
        <div style={{position:"fixed",inset:0,background:"rgba(16,24,40,.4)",backdropFilter:"blur(4px)",zIndex:60,display:"flex",alignItems:"center",justifyContent:"center"}}>
          <div style={{background:T.surface,borderRadius:T.radius,padding:"28px",maxWidth:"320px",width:"90%",textAlign:"center",boxShadow:T.shadowMd,border:`1px solid ${T.border}`,fontFamily:T.font}}>
            <div style={{width:"40px",height:"40px",borderRadius:"10px",background:"#fdecea",display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 14px",fontSize:"18px"}}>🗑</div>
            <h3 style={{margin:"0 0 6px",color:T.text,fontSize:"15px",fontWeight:700}}>Excluir nota fiscal?</h3>
            <p style={{color:T.textMuted,fontSize:"13px",marginBottom:"20px"}}>Essa ação não pode ser desfeita.</p>
            <div style={{display:"flex",gap:"8px"}}>
              <button onClick={()=>setDeleteId(null)} style={{flex:1,padding:"9px",borderRadius:T.radiusSm,border:`1.5px solid ${T.border}`,background:T.bg,color:T.textSub,fontWeight:500,cursor:"pointer",fontFamily:T.font}}>Cancelar</button>
              <button onClick={()=>{setNotas(ns=>ns.filter(n=>n.id!==deleteId));setDeleteId(null);toast("Nota fiscal excluída","error");}} style={{flex:1,padding:"9px",borderRadius:T.radiusSm,border:"none",background:"#F24E29",color:"#fff",fontWeight:600,cursor:"pointer",fontFamily:T.font}}>Excluir</button>
            </div>
          </div>
        </div>
      )}
      <ToastContainer toasts={toasts}/>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════
//  SIDEBAR
// ═══════════════════════════════════════════════════════════════════════════
const NAV = [
  {id:"home",        label:"Dashboard",            icon:"▣"},
  {id:"notas",       label:"Notas Fiscais",         icon:"≡"},
  {id:"boletos",     label:"Boletos Pendentes",     icon:"◷"},
  {id:"fornecedores",label:"Fornecedores",          icon:"◈"},
  {id:"avisos",      label:"Avisos",                icon:"◉"},
];

function Sidebar({ page, setPage, badge, onLogout, userEmail, lastUpdated }) {
  const [, forceUpdate] = React.useState(0);
  React.useEffect(() => {
    const t = setInterval(() => forceUpdate(n=>n+1), 30000);
    return () => clearInterval(t);
  }, []);
  function tempoDesde(d) {
    if (!d) return null;
    const diff = Math.floor((new Date() - d) / 1000);
    if (diff < 60)  return "agora mesmo";
    if (diff < 120) return "há 1 min";
    if (diff < 3600) return `há ${Math.floor(diff/60)} min`;
    return `há ${Math.floor(diff/3600)}h`;
  }
  return (
    <div style={{width:"220px",minWidth:"220px",background:T.surface,borderRight:`1px solid ${T.border}`,display:"flex",flexDirection:"column",height:"100vh",position:"sticky",top:0,fontFamily:T.font}}>
      {/* Logo */}
      <div style={{padding:"24px 20px 20px"}}>
        <div style={{display:"flex",alignItems:"center",gap:"9px"}}>
          <div style={{width:"28px",height:"28px",borderRadius:"7px",background:T.primary,display:"flex",alignItems:"center",justifyContent:"center"}}>
            <span style={{color:"#fff",fontSize:"13px",fontWeight:800}}>K</span>
          </div>
          <div>
            <p style={{margin:0,fontSize:"13.5px",fontWeight:700,color:T.text,lineHeight:1.2}}>Kowalsky</p>
            <p style={{margin:0,fontSize:"11px",color:T.textMuted,lineHeight:1.2}}>Gestão Financeira</p>
          </div>
        </div>
      </div>

      {/* Nav */}
      <nav style={{flex:1,padding:"4px 10px",display:"flex",flexDirection:"column",gap:"1px"}}>
        {NAV.map(n=>{
          const active = page===n.id;
          return (
            <button key={n.id} onClick={()=>setPage(n.id)} style={{
              display:"flex",alignItems:"center",gap:"9px",padding:"8px 10px",
              borderRadius:T.radiusSm,border:"none",cursor:"pointer",
              fontFamily:T.font,fontSize:"13px",fontWeight:active?600:500,
              textAlign:"left",width:"100%",transition:"all .12s",
              background:active?T.primaryLight:"transparent",
              color:active?T.primary:T.textSub,
            }}
            onMouseEnter={e=>{ if(!active) e.currentTarget.style.background=T.bg; }}
            onMouseLeave={e=>{ if(!active) e.currentTarget.style.background="transparent"; }}>
              <span style={{fontSize:"13px",width:"16px",textAlign:"center",opacity:active?1:.7}}>{n.icon}</span>
              <span style={{flex:1}}>{n.label}</span>
              {n.id==="avisos" && badge>0 && null}
            </button>
          );
        })}
      </nav>

      {/* User */}
      <div style={{padding:"12px 16px",borderTop:`1px solid ${T.border}`}}>
        <div style={{marginBottom:"8px",padding:"5px 8px",borderRadius:T.radiusSm,background:T.bg}}>
          <p style={{margin:0,fontSize:"10px",color:T.textMuted,fontFamily:T.mono}}>App atualizado em: 12/03</p>
        </div>
        {lastUpdated && (
          <div style={{display:"flex",alignItems:"center",gap:"5px",marginBottom:"10px",padding:"5px 8px",borderRadius:T.radiusSm,background:T.bg}}>
            <span style={{width:"6px",height:"6px",borderRadius:"50%",background:"#17b26a",flexShrink:0,boxShadow:"0 0 0 2px #dcfce7"}}/>
            <span style={{fontSize:"11px",color:T.textMuted}}>Atualizado {tempoDesde(lastUpdated)}</span>
          </div>
        )}
        <div style={{display:"flex",alignItems:"center",gap:"9px",marginBottom:"8px"}}>
          <div style={{width:"30px",height:"30px",borderRadius:"50%",background:T.primaryLight,display:"flex",alignItems:"center",justifyContent:"center",color:T.primary,fontWeight:700,fontSize:"12px",flexShrink:0,border:`1.5px solid ${T.border}`}}>
            {(userEmail||"?")[0].toUpperCase()}
          </div>
          <div style={{overflow:"hidden",flex:1}}>
            <p style={{margin:0,fontSize:"11px",color:T.textMuted,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{userEmail}</p>
          </div>
        </div>
        <button onClick={onLogout}
          style={{width:"100%",padding:"6px",borderRadius:T.radiusSm,border:`1.5px solid ${T.border}`,background:"transparent",color:T.textMuted,fontSize:"12px",fontWeight:500,cursor:"pointer",fontFamily:T.font,textAlign:"left",paddingLeft:"10px"}}>
          ↩ Sair
        </button>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════
//  BOLETOS NÃO RECEBIDOS
// ═══════════════════════════════════════════════════════════════════════════
function BoletosNaoRecebidosPage({ notas, setNotas }) {
  const [search, setSearch] = useState("");
  const [filtro, setFiltro] = useState("todos");
  const { toasts, toast } = useToast();
  const [confirmId, setConfirmId] = useState(null); // id da NF aguardando confirmação

  const td  = today();
  const ew  = endOfWeek();
  const p30 = (() => { const d = new Date(); d.setDate(d.getDate()+30); return d.toISOString().split("T")[0]; })();

  // 1 linha por NF — apenas as que NÃO tiveram boleto recebido
  const nfsPendentes = useMemo(() => {
    return notas
      .filter(n => !n.boletosRecebidos)
      .map(n => {
        const parcelasAbertas = n.vencimentos.filter(v => parcelaStatusEfetivo(n,v) !== "pago");
        const proximoVenc = [...parcelasAbertas].sort()[0] || n.vencimentos[0];
        const valorRestante = parcelasAbertas.length * (n.valor / n.parcelas);
        return { ...n, parcelasAbertas, proximoVenc, valorRestante };
      })
      .sort((a,b) => a.proximoVenc.localeCompare(b.proximoVenc));
  }, [notas]);

  const filtered = useMemo(() => {
    return nfsPendentes.filter(n => {
      const ms = !search || n.numero.toLowerCase().includes(search.toLowerCase()) || n.fornecedor.toLowerCase().includes(search.toLowerCase());
      if (!ms) return false;
      if (filtro === "vencidos") return n.proximoVenc < td;
      if (filtro === "semana")   return n.proximoVenc >= td && n.proximoVenc <= ew;
      if (filtro === "30dias")   return n.proximoVenc >= td && n.proximoVenc <= p30;
      return true;
    });
  }, [nfsPendentes, search, filtro, td, ew, p30]);

  const counts = {
    todos:    nfsPendentes.length,
    vencidos: nfsPendentes.filter(n => n.proximoVenc < td).length,
    semana:   nfsPendentes.filter(n => n.proximoVenc >= td && n.proximoVenc <= ew).length,
    p30dias:  nfsPendentes.filter(n => n.proximoVenc >= td && n.proximoVenc <= p30).length,
  };

  const [sortCol, setSortCol] = useState("proximoVenc");
  const [sortDir, setSortDir] = useState("asc");

  function toggleSort(col) {
    setSortCol(prev => { if(prev===col){setSortDir(d=>d==="asc"?"desc":"asc");return col;} setSortDir("asc");return col; });
  }

  const sorted = useMemo(() => {
    return [...filtered].sort((a,b) => {
      let va, vb;
      if      (sortCol==="numero")       { va=a.numero;       vb=b.numero; }
      else if (sortCol==="fornecedor")   { va=a.fornecedor;   vb=b.fornecedor; }
      else if (sortCol==="empresa")      { va=a.empresa;      vb=b.empresa; }
      else if (sortCol==="proximoVenc")  { va=a.proximoVenc;  vb=b.proximoVenc; }
      else if (sortCol==="parcelas")     { return sortDir==="asc"?a.parcelasAbertas.length-b.parcelasAbertas.length:b.parcelasAbertas.length-a.parcelasAbertas.length; }
      else if (sortCol==="valor")        { return sortDir==="asc"?a.valorRestante-b.valorRestante:b.valorRestante-a.valorRestante; }
      else return 0;
      return sortDir==="asc"?va.localeCompare(vb,"pt-BR"):vb.localeCompare(va,"pt-BR");
    });
  }, [filtered, sortCol, sortDir]);

  function marcarBoletoRecebido(id) {
    setNotas(ns => ns.map(n => n.id===id ? {...n, boletosRecebidos:true} : n));
    toast("Boleto marcado como recebido");
  }

  function urgencia(venc) {
    if (venc < td)   return { color:"#b42318", bg:"#fdecea",      label:"Vencido" };
    if (venc === td) return { color:"#b54708", bg:"#fffaeb",      label:"Vence hoje" };
    if (venc <= ew)  return { color:T.primary, bg:T.primaryLight, label:"Esta semana" };
    return null;
  }

  return (
    <div style={{padding:"36px 40px", fontFamily:T.font, maxWidth:"1200px"}}>
      <div style={{marginBottom:"24px"}}>
        <h1 style={{margin:0, fontSize:"22px", fontWeight:700, color:T.text, letterSpacing:"-.02em"}}>Boletos Pendentes</h1>
        <p style={{margin:"4px 0 0", fontSize:"13px", color:T.textMuted}}>Notas fiscais com boleto físico ainda não recebido</p>
      </div>

      {counts.vencidos > 0 && (
        <div style={{background:T.dangerLight, border:"1px solid #fca5a5", borderRadius:T.radiusSm, padding:"10px 14px", marginBottom:"16px", fontSize:"13px", color:"#b42318", display:"flex", alignItems:"center", gap:"8px", fontWeight:500}}>
          ⚠ <strong>{counts.vencidos} NF(s)</strong> com próximo vencimento já passado e boleto não recebido.
        </div>
      )}

      <div style={{marginBottom:"12px"}}>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar por número da NF ou fornecedor..." style={{...S.input, maxWidth:"400px", background:T.surface, boxShadow:T.shadow}}/>
      </div>

      <div style={{display:"flex", gap:"6px", marginBottom:"18px", flexWrap:"wrap"}}>
        {[
          ["todos",    `Todos (${counts.todos})`],
          ["vencidos", `Vencidos (${counts.vencidos})`],
          ["semana",   `Esta Semana (${counts.semana})`],
          ["30dias",   `Próximos 30 Dias (${counts.p30dias})`],
        ].map(([k,l])=>(
          <button key={k} onClick={()=>setFiltro(k)} style={{padding:"6px 14px", borderRadius:"20px", cursor:"pointer", fontSize:"12.5px", fontWeight:600, border:`1.5px solid ${filtro===k?T.primary:T.border}`, background:filtro===k?T.primary:T.surface, color:filtro===k?"#fff":T.textSub, transition:"all .15s", fontFamily:T.font}}>{l}</button>
        ))}
      </div>

      <div style={{background:T.surface, borderRadius:T.radius, boxShadow:T.shadow, border:`1px solid ${T.border}`, overflow:"hidden"}}>
        <table style={{width:"100%", borderCollapse:"collapse", fontSize:"13px", fontFamily:T.font}}>
          <thead>
            <tr style={{background:T.bg}}>
              {[
                {label:"Número NF",        col:"numero"},
                {label:"Fornecedor",        col:"fornecedor"},
                {label:"Empresa",           col:"empresa"},
                {label:"Próx. Vencimento",  col:"proximoVenc"},
                {label:"Parcelas Abertas",  col:"parcelas"},
                {label:"Valor em Aberto",   col:"valor"},
                {label:"Situação",          col:null},
                {label:"Ações",             col:null},
              ].map(({label,col})=>(
                <th key={label} onClick={col?()=>toggleSort(col):undefined}
                  style={{padding:"10px 16px", textAlign:"left", fontWeight:600, color:col&&sortCol===col?T.primary:T.textMuted, fontSize:"11px", textTransform:"uppercase", letterSpacing:".06em", borderBottom:`1px solid ${T.border}`, whiteSpace:"nowrap", cursor:col?"pointer":"default", userSelect:"none"}}>
                  {label}{col&&sortCol===col?(sortDir==="asc"?" ▲":" ▼"):(col?" ↕":"")}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.length===0 ? (
              <tr><td colSpan={8} style={{padding:"48px", textAlign:"center", color:T.textMuted, fontSize:"13px"}}>
                {nfsPendentes.length===0 ? "✓ Todos os boletos foram recebidos" : "Nenhum resultado para este filtro"}
              </td></tr>
            ) : sorted.map((n,i)=>{
              const urg = urgencia(n.proximoVenc);
              const isLast = i === sorted.length - 1;
              return (
                <tr key={n.id}
                  style={{borderBottom:isLast?"none":`1px solid ${T.border}`, transition:"background .1s"}}
                  onMouseEnter={e=>e.currentTarget.style.background=T.bg}
                  onMouseLeave={e=>e.currentTarget.style.background=""}>
                  <td style={{padding:"13px 16px", fontFamily:T.mono, fontWeight:600, color:T.text, fontSize:"12.5px"}}>{n.numero}</td>
                  <td style={{padding:"13px 16px", color:T.text, fontWeight:500}}>{fnome(n.fornecedor)}</td>
                  <td style={{padding:"13px 16px"}}>
                    <span style={{background:T.bg, color:T.textSub, borderRadius:"5px", padding:"2px 8px", fontSize:"11.5px", fontWeight:500, border:`1px solid ${T.border}`}}>{n.empresa}</span>
                  </td>
                  <td style={{padding:"13px 16px", fontWeight:urg?600:400, color:urg?.color||T.textSub, fontFamily:T.mono, fontSize:"12.5px"}}>
                    {fdate(n.proximoVenc)}
                  </td>
                  <td style={{padding:"13px 16px", color:T.textSub, fontSize:"12.5px"}}>
                    {n.parcelasAbertas.length}/{n.parcelas}
                  </td>
                  <td style={{padding:"13px 16px", fontWeight:600, color:T.text, fontFamily:T.mono, fontSize:"12.5px"}}>
                    {fmt(n.valorRestante)}
                  </td>
                  <td style={{padding:"13px 16px"}}>
                    {urg
                      ? <span style={{fontSize:"11.5px", fontWeight:600, padding:"2px 9px", borderRadius:"20px", background:urg.bg, color:urg.color}}>{urg.label}</span>
                      : <span style={{fontSize:"12px", color:T.textMuted}}>—</span>
                    }
                  </td>
                  <td style={{padding:"13px 16px"}}>
                    <button onClick={()=>setConfirmId(n.id)}
                      style={{padding:"5px 12px", borderRadius:T.radiusSm, border:`1.5px solid ${T.border}`, background:T.surface, color:T.primary, fontWeight:600, fontSize:"12px", cursor:"pointer", fontFamily:T.font, whiteSpace:"nowrap"}}>
                      ✓ Boleto Recebido
                    </button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      {/* Modal de confirmação */}
      {confirmId !== null && (() => {
        const nf = notas.find(n => n.id === confirmId);
        return (
          <div style={{position:"fixed",inset:0,background:"rgba(1,21,38,.45)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:T.font}}>
            <div style={{background:T.surface,borderRadius:T.radius,boxShadow:T.shadowMd,padding:"28px 32px",maxWidth:"400px",width:"90%",border:`1px solid ${T.border}`}}>
              <div style={{width:"44px",height:"44px",borderRadius:"50%",background:T.primaryLight,display:"flex",alignItems:"center",justifyContent:"center",fontSize:"22px",marginBottom:"16px"}}>✓</div>
              <h3 style={{margin:"0 0 6px",fontSize:"16px",fontWeight:700,color:T.text}}>Confirmar recebimento</h3>
              <p style={{margin:"0 0 6px",fontSize:"13.5px",color:T.textSub}}>
                Você confirma que o boleto da nota <strong style={{color:T.text,fontFamily:T.mono}}>{nf?.numero}</strong> foi fisicamente recebido?
              </p>
              <p style={{margin:"0 0 22px",fontSize:"12px",color:T.textMuted}}>{fnome(nf?.fornecedor)} · {nf?.empresa}</p>
              <div style={{display:"flex",gap:"10px"}}>
                <button onClick={()=>setConfirmId(null)}
                  style={{flex:1,padding:"10px",borderRadius:T.radiusSm,border:`1.5px solid ${T.border}`,background:T.bg,color:T.textSub,fontWeight:600,fontSize:"13px",cursor:"pointer",fontFamily:T.font}}>
                  Cancelar
                </button>
                <button onClick={()=>{ marcarBoletoRecebido(confirmId); setConfirmId(null); }}
                  style={{flex:2,padding:"10px",borderRadius:T.radiusSm,border:"none",background:T.primary,color:"#fff",fontWeight:700,fontSize:"13px",cursor:"pointer",fontFamily:T.font}}>
                  Sim, boleto recebido
                </button>
              </div>
            </div>
          </div>
        );
      })()}
      <ToastContainer toasts={toasts}/>
    </div>
  );
}


function PlaceholderPage({title,icon}) {
  return (
    <div style={{padding:"32px 36px",fontFamily:T.font,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",minHeight:"60vh",color:T.textMuted,textAlign:"center"}}>
      <div style={{fontSize:"48px",marginBottom:"16px"}}>{icon}</div>
      <h2 style={{margin:0,fontSize:"20px",fontWeight:700,color:"#cbd5e1"}}>{title}</h2>
      <p style={{margin:"8px 0 0",fontSize:"14px"}}>Em breve disponível</p>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════
//  TAREFAS — helpers
// ═══════════════════════════════════════════════════════════════════════════
const RECORRENCIA_CFG = {
  uma_vez:  { label:"Uma vez",      color:"#475569", bg:"#f1f5f9" },
  diaria:   { label:"Diária",       color:"#1A5173", bg:"#C4DDF2" },
  semanal:  { label:"Semanal",      color:"#7c3aed", bg:"#f5f3ff" },
  mensal:   { label:"Mensal",       color:"#0891b2", bg:"#ecfeff" },
  anual:    { label:"Anual",        color:"#059669", bg:"#ecfdf5" },
};

const DIAS_SEMANA = ["Domingo","Segunda","Terça","Quarta","Quinta","Sexta","Sábado"];

// Retorna o status efetivo de uma parcela:
// Prioridade: manual > automático (boleto recebido + vencido)
function parcelaStatusEfetivo(nota, venc) {
  if (nota.parcelasPagas && nota.parcelasPagas.includes(venc)) return "pago";
  if (nota.boletosRecebidos && venc < today()) return "pago";
  return "pendente";
}

// Status calculado automaticamente — nunca manual
function statusNota(n) {
  if (!n || !n.vencimentos || n.vencimentos.length === 0) return { key:"incompleto", label:"Incompleto", color:"#b42318", bg:"#fdecea" };
  const td = new Date().toISOString().split("T")[0];
  const todasPagas = n.vencimentos.every(v => parcelaStatusEfetivo(n, v) === "pago");
  if (todasPagas)          return { key:"quitado",    label:"Quitado",    color:"#166534", bg:"#dcfce7" };
  const temVencida = n.vencimentos.some(v => v < td && parcelaStatusEfetivo(n, v) !== "pago");
  if (temVencida)          return { key:"vencido",    label:"Vencido ⚠",  color:"#92400e", bg:"#fef3c7" };
  if (!n.boletosRecebidos) return { key:"incompleto", label:"Incompleto", color:"#b42318", bg:"#fdecea" };
  return                          { key:"aberto",     label:"Aberto",     color:"#1547a0", bg:"#dbeafe" };
}

// Retorna true se TODAS as parcelas de uma NF já foram pagas (manual ou automático)
function notaPagaAutomaticamente(nota) {
  return nota.vencimentos.every(v => parcelaStatusEfetivo(nota, v) === "pago");
}

function tarefaAtivaHoje(t) {
  const now = new Date();
  const td  = today();
  if (t.recorrencia === "uma_vez")  return t.data === td;
  if (t.recorrencia === "diaria")   return true;
  if (t.recorrencia === "semanal")  return now.getDay() === (t.diaSemana ?? new Date(t.data+"T12:00:00").getDay());
  if (t.recorrencia === "mensal")   return now.getDate() === (t.diaMes ?? new Date(t.data+"T12:00:00").getDate());
  if (t.recorrencia === "anual") {
    const d = new Date(t.data+"T12:00:00");
    return now.getDate()===d.getDate() && now.getMonth()===d.getMonth();
  }
  return false;
}

// Próxima ocorrência (texto amigável)
function proximaOcorrencia(t) {
  if (t.recorrencia === "uma_vez")  return `Uma vez em ${fdate(t.data)}`;
  if (t.recorrencia === "diaria")   return "Todo dia";
  if (t.recorrencia === "semanal") {
    const dia = t.diaSemana ?? new Date(t.data+"T12:00:00").getDay();
    return `Toda ${DIAS_SEMANA[dia]}`;
  }
  if (t.recorrencia === "mensal") {
    const d = t.diaMes ?? new Date(t.data+"T12:00:00").getDate();
    return `Todo dia ${d} do mês`;
  }
  if (t.recorrencia === "anual") {
    const d = new Date(t.data+"T12:00:00");
    const mes = d.toLocaleString("pt-BR",{month:"long"});
    return `Todo ano em ${d.getDate()} de ${mes}`;
  }
  return "";
}

const TAREFAS_SEED = [
  { id:1, titulo:"Verificar boletos do dia",    recorrencia:"diaria",  data:today(), diaSemana:null, diaMes:null, concluidoEm:null },
  { id:2, titulo:"Enviar relatório semanal",     recorrencia:"semanal", data:today(), diaSemana:5,   diaMes:null, concluidoEm:null },
  { id:3, titulo:"Fechar fluxo de caixa mensal", recorrencia:"mensal",  data:today(), diaSemana:null, diaMes:1,   concluidoEm:null },
];

// ═══════════════════════════════════════════════════════════════════════════
//  MINI CALENDAR COMPONENT
// ═══════════════════════════════════════════════════════════════════════════
function MiniCalendar({ value, onChange }) {
  const sel   = value ? new Date(value + "T12:00:00") : new Date();
  const [view, setView] = useState({ year: sel.getFullYear(), month: sel.getMonth() });

  const firstDay  = new Date(view.year, view.month, 1).getDay();
  const daysInMonth = new Date(view.year, view.month + 1, 0).getDate();
  const todayStr  = today();

  const MESES = ["janeiro","fevereiro","março","abril","maio","junho","julho","agosto","setembro","outubro","novembro","dezembro"];
  const DIAS_H = ["dom","seg","ter","qua","qui","sex","sáb"];

  function prevMonth() {
    setView(v => v.month === 0 ? {year:v.year-1,month:11} : {year:v.year,month:v.month-1});
  }
  function nextMonth() {
    setView(v => v.month === 11 ? {year:v.year+1,month:0} : {year:v.year,month:v.month+1});
  }
  function selectDay(d) {
    const m = String(view.month+1).padStart(2,"0");
    const dd = String(d).padStart(2,"0");
    onChange(`${view.year}-${m}-${dd}`);
  }

  const cells = [];
  for (let i=0;i<firstDay;i++) cells.push(null);
  for (let d=1;d<=daysInMonth;d++) cells.push(d);

  const selStr = value;
  function cellStr(d) {
    return `${view.year}-${String(view.month+1).padStart(2,"0")}-${String(d).padStart(2,"0")}`;
  }

  return (
    <div style={{border:"1.5px solid #e2e8f0",borderRadius:"10px",background:T.surface,padding:"12px",userSelect:"none"}}>
      {/* Header */}
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:"10px"}}>
        <button onClick={prevMonth} style={{border:"none",background:"none",cursor:"pointer",fontSize:"16px",color:"#475569",padding:"2px 6px"}}>‹</button>
        <span style={{fontSize:"13px",fontWeight:700,color:T.text,textTransform:"capitalize"}}>
          {MESES[view.month]} de {view.year}
        </span>
        <button onClick={nextMonth} style={{border:"none",background:"none",cursor:"pointer",fontSize:"16px",color:"#475569",padding:"2px 6px"}}>›</button>
      </div>
      {/* Dias da semana */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:"2px",marginBottom:"4px"}}>
        {DIAS_H.map(d=>(
          <div key={d} style={{textAlign:"center",fontSize:"11px",fontWeight:700,color:T.textMuted,padding:"3px 0"}}>{d}</div>
        ))}
      </div>
      {/* Células */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:"2px"}}>
        {cells.map((d,i)=>{
          if (!d) return <div key={`e${i}`}/>;
          const cs = cellStr(d);
          const isToday = cs===todayStr;
          const isSel   = cs===selStr;
          return (
            <button key={d} onClick={()=>selectDay(d)} style={{
              border:"none",borderRadius:"6px",padding:"5px 0",fontSize:"12.5px",fontWeight:isSel||isToday?700:400,
              cursor:"pointer",textAlign:"center",transition:"background .1s",
              background: isSel?"#1A5173":isToday?"#C4DDF2":"transparent",
              color: isSel?"#fff":isToday?"#1A5173":"#374151",
            }}>{d}</button>
          );
        })}
      </div>
      {/* Rodapé hoje */}
      <div style={{marginTop:"8px",paddingTop:"8px",borderTop:"1px solid #f1f5f9",textAlign:"center"}}>
        <button onClick={()=>onChange(todayStr)} style={{border:"none",background:"none",cursor:"pointer",fontSize:"11.5px",color:"#1A5173",fontWeight:600}}>
          Hoje: {fdate(todayStr)}
        </button>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════
//  MODAL DE TAREFA
// ═══════════════════════════════════════════════════════════════════════════
const REPETIR_OPTS = [
  { value:"uma_vez",    label:"Uma vez"      },
  { value:"diaria",     label:"Diariamente"  },
  { value:"semanal",    label:"Semanalmente" },
  { value:"mensal",     label:"Mensalmente"  },
  { value:"anual",      label:"Anualmente"   },
];

function TarefaModal({ onClose, onSave, editData }) {
  const blank = { titulo:"", descricao:"", recorrencia:"diaria", data:today(), diaSemana:null, diaMes:null };
  const [form, setForm] = useState(editData || blank);
  const [errors, setErrors] = useState({});
  const set = (f,v) => { setForm(p=>({...p,[f]:v})); setErrors(p=>({...p,[f]:undefined})); };

  function handleDataChange(v) {
    const d = new Date(v + "T12:00:00");
    setForm(p=>({...p, data:v, diaSemana:d.getDay(), diaMes:d.getDate()}));
    setErrors(p=>({...p, data:undefined}));
  }

  function handleSave() {
    const e={};
    if (!form.titulo.trim()) e.titulo="Digite um título para a tarefa";
    if (!form.data)          e.data="Selecione uma data";
    setErrors(e);
    if (Object.keys(e).length) return;
    const d = new Date(form.data + "T12:00:00");
    onSave({ ...form, diaSemana:d.getDay(), diaMes:d.getDate(), id:editData?.id||Date.now(), concluidoEm:editData?.concluidoEm||null });
  }

  return (
    <div style={{position:"fixed",inset:0,background:"rgba(15,23,42,0.5)",backdropFilter:"blur(4px)",zIndex:60,display:"flex",alignItems:"center",justifyContent:"center",padding:"16px"}}>
      <div style={{background:T.surface,borderRadius:"16px",width:"100%",maxWidth:"460px",maxHeight:"92vh",overflowY:"auto",boxShadow:"0 24px 64px rgba(0,0,0,0.18)",fontFamily:T.font}}>
        <div style={{padding:"22px 26px 0",display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
          <div>
            <h2 style={{margin:0,fontSize:"18px",fontWeight:800,color:T.text}}>{editData?"Editar Tarefa":"Nova Tarefa"}</h2>
            <p style={{margin:"3px 0 0",fontSize:"12.5px",color:T.textSub}}>Configure os detalhes da tarefa</p>
          </div>
          <button onClick={onClose} style={{border:"none",background:"none",cursor:"pointer",fontSize:"22px",color:T.textMuted}}>×</button>
        </div>

        <div style={{padding:"18px 26px 26px",display:"flex",flexDirection:"column",gap:"16px"}}>

          {/* Título */}
          <div>
            <label style={S.label}>Título <Red/></label>
            <input value={form.titulo} onChange={e=>set("titulo",e.target.value)} placeholder="Ex: Verificar boletos do dia" style={{...S.input,borderColor:errors.titulo?"#F24E29":"#e2e8f0"}}/>
            {errors.titulo&&<span style={S.error}>{errors.titulo}</span>}
          </div>

          {/* Descrição */}
          <div>
            <label style={S.label}>Descrição:</label>
            <textarea
              value={form.descricao}
              onChange={e=>set("descricao",e.target.value)}
              rows={4}
              style={{...S.input,resize:"vertical",minHeight:"80px",borderColor:"#e2e8f0"}}
            />
          </div>

          {/* Repetir */}
          <div>
            <label style={S.label}>Repetir:</label>
            <div style={{position:"relative"}}>
              <select
                value={form.recorrencia}
                onChange={e=>set("recorrencia",e.target.value)}
                style={{...S.input,appearance:"none",paddingRight:"36px",cursor:"pointer",background:T.surface}}>
                {REPETIR_OPTS.map(o=><option key={o.value} value={o.value}>{o.label}</option>)}
              </select>
              <span style={{position:"absolute",right:"12px",top:"50%",transform:"translateY(-50%)",color:T.textSub,pointerEvents:"none",fontSize:"12px"}}>▼</span>
            </div>
          </div>

          {/* Data */}
          <div>
            <label style={S.label}>Data:</label>
            <MiniCalendar value={form.data} onChange={handleDataChange}/>
            {errors.data&&<span style={S.error}>{errors.data}</span>}
          </div>

          <div style={{display:"flex",gap:"10px",marginTop:"4px"}}>
            <button onClick={onClose} style={{flex:1,padding:"11px",borderRadius:"9px",border:"1.5px solid #e2e8f0",background:T.bg,color:"#475569",fontWeight:600,fontSize:"13.5px",cursor:"pointer"}}>Cancelar</button>
            <button onClick={handleSave} style={{flex:2,padding:"11px",borderRadius:"9px",border:"none",background:"linear-gradient(135deg,#1A5173,#1A5173)",color:"#fff",fontWeight:700,fontSize:"13.5px",cursor:"pointer",boxShadow:"0 4px 14px rgba(37,99,235,.3)"}}>{editData?"Salvar":"Criar Tarefa"}</button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════
//  AVISOS PAGE
// ═══════════════════════════════════════════════════════════════════════════
function AvisosPage({ tarefas, setTarefas }) {
  const [showModal, setShowModal] = useState(false);
  const [editTarefa, setEditTarefa] = useState(null);
  const [deleteId, setDeleteId] = useState(null);
  const [filtro, setFiltro] = useState("todas");

  function handleSave(t) {
    setTarefas(ts => t.id && ts.find(x=>x.id===t.id) ? ts.map(x=>x.id===t.id?t:x) : [...ts,t]);
    setShowModal(false); setEditTarefa(null);
  }

  function concluir(id) {
    setTarefas(ts=>ts.map(t=>t.id===id?{...t,concluidoEm:today()}:t));
  }

  function resetar(id) {
    setTarefas(ts=>ts.map(t=>t.id===id?{...t,concluidoEm:null}:t));
  }

  const pendentesHoje = tarefas.filter(t=>tarefaAtivaHoje(t)&&t.concluidoEm!==today()).length;

  const filtradas = tarefas.filter(t => {
    if (filtro==="hoje")      return tarefaAtivaHoje(t) && t.concluidoEm!==today();
    if (filtro==="concluidas") return t.concluidoEm===today();
    return true;
  });

  return (
    <div style={{padding:"28px 36px",fontFamily:T.font}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:"24px"}}>
        <div>
          <h1 style={{margin:0,fontSize:"22px",fontWeight:800,color:T.text}}>Avisos & Tarefas</h1>
          <p style={{margin:"3px 0 0",fontSize:"13px",color:T.textSub}}>Tarefas recorrentes do seu dia a dia</p>
        </div>
        <button onClick={()=>setShowModal(true)} style={{padding:"9px 18px",borderRadius:T.radiusSm,border:"none",background:"linear-gradient(135deg,#1A5173,#1A5173)",color:"#fff",fontWeight:700,fontSize:"13px",cursor:"pointer",boxShadow:"0 4px 12px rgba(37,99,235,.3)"}}>+ Nova Tarefa</button>
      </div>

      {/* Cards resumo */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:"14px",marginBottom:"24px"}}>
        {[
          {label:"Pendentes Hoje", v:pendentesHoje, icon:"⏰", c:"#d97706", bg:"#fffbeb"},
          {label:"Tarefas Diárias", v:tarefas.filter(t=>t.recorrencia==="diaria").length, icon:"📅", c:"#1A5173", bg:"#C4DDF2"},
          {label:"Total de Tarefas", v:tarefas.length, icon:"📋", c:"#7c3aed", bg:"#f5f3ff"},
        ].map(c=>(
          <div key={c.label} style={{background:T.surface,borderRadius:T.radius,padding:"16px 18px",boxShadow:"0 1px 3px rgba(0,0,0,.05)",border:`1px solid ${T.border}`}}>
            <div style={{display:"flex",alignItems:"center",gap:"8px",marginBottom:"8px"}}>
              <div style={{width:"30px",height:"30px",borderRadius:"7px",background:c.bg,display:"flex",alignItems:"center",justifyContent:"center",fontSize:"15px"}}>{c.icon}</div>
              <span style={{fontSize:"12px",color:T.textSub,fontWeight:500}}>{c.label}</span>
            </div>
            <div style={{fontSize:"22px",fontWeight:800,color:T.text}}>{c.v}</div>
          </div>
        ))}
      </div>

      {/* Filtros */}
      <div style={{display:"flex",gap:"6px",marginBottom:"16px"}}>
        {[["todas","Todas"],["hoje","Pendentes Hoje"],["concluidas","Concluídas Hoje"]].map(([k,l])=>(
          <button key={k} onClick={()=>setFiltro(k)} style={{padding:"6px 14px",borderRadius:"20px",cursor:"pointer",fontSize:"12.5px",fontWeight:600,border:filtro===k?"none":"1.5px solid #e2e8f0",background:filtro===k?"#1A5173":"#fff",color:filtro===k?"#fff":"#475569",transition:"all .15s"}}>{l}</button>
        ))}
      </div>

      {/* Lista de tarefas */}
      <div style={{display:"flex",flexDirection:"column",gap:"10px"}}>
        {filtradas.length===0 ? (
          <div style={{background:T.surface,borderRadius:T.radius,padding:"40px",textAlign:"center",color:T.textMuted,border:`1px solid ${T.border}`}}>
            Nenhuma tarefa encontrada
          </div>
        ) : filtradas.map(t=>{
          const rec = RECORRENCIA_CFG[t.recorrencia];
          const ativa = tarefaAtivaHoje(t);
          const concluida = t.concluidoEm===today();
          return (
            <div key={t.id} style={{background:T.surface,borderRadius:T.radius,padding:"16px 20px",boxShadow:"0 1px 3px rgba(0,0,0,.05)",border:`1px solid ${concluida?"#d1fae5":ativa?"#e0e7ff":"#f1f5f9"}`,display:"flex",alignItems:"center",gap:"14px",transition:"all .2s"}}>
              {/* Checkbox */}
              <button
                onClick={()=>concluida?resetar(t.id):concluir(t.id)}
                style={{width:"22px",height:"22px",borderRadius:"50%",border:`2px solid ${concluida?"#10b981":"#cbd5e1"}`,background:concluida?"#10b981":"#fff",cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,transition:"all .2s"}}>
                {concluida&&<span style={{color:"#fff",fontSize:"12px",fontWeight:700}}>✓</span>}
              </button>

              <div style={{flex:1}}>
                <p style={{margin:0,fontSize:"14px",fontWeight:600,color:concluida?T.textMuted:T.text,textDecoration:concluida?"line-through":"none"}}>{t.titulo}</p>
                <p style={{margin:"3px 0 0",fontSize:"12px",color:T.textMuted}}>{proximaOcorrencia(t)}</p>
              </div>

              <div style={{display:"flex",alignItems:"center",gap:"8px"}}>
                <span style={{fontSize:"11.5px",fontWeight:600,padding:"3px 10px",borderRadius:"20px",background:rec.bg,color:rec.color}}>{rec.label}</span>
                {ativa&&!concluida&&<span style={{fontSize:"11.5px",fontWeight:600,padding:"3px 10px",borderRadius:"20px",background:"#fef3c7",color:"#d97706"}}>Hoje</span>}
                {concluida&&<span style={{fontSize:"11.5px",fontWeight:600,padding:"3px 10px",borderRadius:"20px",background:"#d1fae5",color:"#065f46"}}>✓ Feita</span>}
              </div>

              <div style={{display:"flex",gap:"5px"}}>
                <button onClick={()=>{setEditTarefa(t);}} style={{padding:"5px 9px",borderRadius:"6px",border:"1.5px solid #e2e8f0",background:T.bg,cursor:"pointer",fontSize:"12px"}}>✏️</button>
                <button onClick={()=>setDeleteId(t.id)} style={{padding:"5px 9px",borderRadius:"6px",border:"1.5px solid #fee2e2",background:"#fff5f5",cursor:"pointer",fontSize:"12px"}}>🗑</button>
              </div>
            </div>
          );
        })}
      </div>

      {(showModal||editTarefa)&&<TarefaModal onClose={()=>{setShowModal(false);setEditTarefa(null);}} onSave={handleSave} editData={editTarefa}/>}

      {deleteId&&(
        <div style={{position:"fixed",inset:0,background:"rgba(15,23,42,.5)",backdropFilter:"blur(4px)",zIndex:60,display:"flex",alignItems:"center",justifyContent:"center"}}>
          <div style={{background:T.surface,borderRadius:T.radius,padding:"28px",maxWidth:"340px",width:"90%",textAlign:"center",boxShadow:"0 24px 60px rgba(0,0,0,.18)"}}>
            <div style={{fontSize:"34px",marginBottom:"10px"}}>🗑️</div>
            <h3 style={{margin:"0 0 8px",color:T.text}}>Excluir tarefa?</h3>
            <p style={{color:T.textSub,fontSize:"13.5px",marginBottom:"20px"}}>Essa ação não pode ser desfeita.</p>
            <div style={{display:"flex",gap:"10px"}}>
              <button onClick={()=>setDeleteId(null)} style={{flex:1,padding:"10px",borderRadius:T.radiusSm,border:"1.5px solid #e2e8f0",background:T.bg,color:"#475569",fontWeight:600,cursor:"pointer"}}>Cancelar</button>
              <button onClick={()=>{setTarefas(ts=>ts.filter(t=>t.id!==deleteId));setDeleteId(null);}} style={{flex:1,padding:"10px",borderRadius:T.radiusSm,border:"none",background:"#F24E29",color:"#fff",fontWeight:700,cursor:"pointer"}}>Excluir</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════
//  RELATÓRIO POR FORNECEDOR
// ═══════════════════════════════════════════════════════════════════════════
function RelatorioFornecedor({ fornecedor, notas, onVoltar, lastAddedId }) {
  const [aba, setAba]           = useState("notas");
  const [expandedId, setExpandedId] = useState(null);
  const [fDataDe, setFDataDe]   = useState("");
  const [fDataAte, setFDataAte] = useState("");
  const [ordenar, setOrdenar]   = useState("recente"); // recente | maior | menor
  const td = today();

  const notasForn = notas.filter(n => n.fornecedorId === fornecedor.id);

  // Filtrar por período
  const notasFiltradas = notasForn.filter(n => {
    if (fDataDe  && n.emissao < fDataDe)  return false;
    if (fDataAte && n.emissao > fDataAte) return false;
    return true;
  });

  // Ordenar
  const notasOrdenadas = [...notasFiltradas].sort((a, b) => {
    if (ordenar === "recente") return b.emissao.localeCompare(a.emissao);
    if (ordenar === "maior")   return b.valor - a.valor;
    if (ordenar === "menor")   return a.valor - b.valor;
    return 0;
  });

  const totalGeral    = notasFiltradas.reduce((s,n)=>s+n.valor,0);
  const totalPago     = notasFiltradas.filter(n=>statusNota(n).key==="quitado").reduce((s,n)=>s+n.valor,0);
  const totalPendente = notasFiltradas.filter(n=>statusNota(n).key!=="quitado").reduce((s,n)=>s+n.valor,0);
  const totalParcelas = notasFiltradas.reduce((s,n)=>s+n.parcelas,0);

  const porEmpresa = EMPRESAS.map((emp,i)=>({
    name:emp, color:["#F24E29","#5B89A6","#1A5173"][i],
    total:notasFiltradas.filter(n=>n.empresa===emp).reduce((s,n)=>s+n.valor,0),
    count:notasFiltradas.filter(n=>n.empresa===emp).length,
  }));
  const maxEmp = Math.max(...porEmpresa.map(e=>e.total),1);

  // Próximos vencimentos
  const parcelas = [];
  notasFiltradas.forEach(n => {
    n.vencimentos.forEach((v,i)=>{
      if(v >= td) parcelas.push({ notaId:n.id, numero:n.numero, empresa:n.empresa, venc:v, parcela:i+1, total:n.parcelas, valor:n.valor/n.parcelas, pago: parcelaStatusEfetivo(n,v)==="pago" });
    });
  });
  const proximos = parcelas.sort((a,b)=>a.venc.localeCompare(b.venc)).slice(0,10);

  const temFiltro = fDataDe || fDataAte;

  return (
    <div style={{padding:"28px 36px",fontFamily:T.font}}>

      {/* Cabeçalho */}
      <div style={{background:T.surface,borderRadius:T.radius,padding:"20px 24px",marginBottom:"20px",boxShadow:"0 1px 4px rgba(0,0,0,.06)",border:`1px solid ${T.border}`,display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:"12px"}}>
        <div style={{display:"flex",alignItems:"center",gap:"14px"}}>
          <button onClick={onVoltar} style={{padding:"7px 14px",borderRadius:T.radiusSm,border:"1.5px solid #e2e8f0",background:T.bg,color:"#475569",fontWeight:600,fontSize:"13px",cursor:"pointer"}}>
            ← Voltar
          </button>
          <div style={{width:"1px",height:"32px",background:"#e2e8f0"}}/>
          <div style={{width:"36px",height:"36px",borderRadius:"9px",background:"#C4DDF2",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"18px"}}>🏢</div>
          <div>
            <h1 style={{margin:0,fontSize:"18px",fontWeight:800,color:T.text}}>{fornecedor.nome}</h1>
            <p style={{margin:0,fontSize:"12px",color:T.textMuted}}>{notasForn.length} nota(s) no total</p>
          </div>
        </div>

        {/* Filtros de data + ordenação */}
        <div style={{display:"flex",alignItems:"center",gap:"8px",flexWrap:"wrap"}}>
          {/* Período */}
          <div style={{display:"flex",alignItems:"center",gap:"6px",background:T.bg,padding:"6px 10px",borderRadius:T.radiusSm,border:`1px solid ${T.border}`}}>
            <span style={{fontSize:"12px",color:T.textMuted,fontWeight:500}}>De</span>
            <input type="date" value={fDataDe} onChange={e=>setFDataDe(e.target.value)}
              style={{border:"none",background:"transparent",fontSize:"12.5px",color:T.text,outline:"none",fontFamily:T.font,cursor:"pointer"}}/>
            <span style={{fontSize:"12px",color:T.textMuted,fontWeight:500}}>até</span>
            <input type="date" value={fDataAte} onChange={e=>setFDataAte(e.target.value)}
              style={{border:"none",background:"transparent",fontSize:"12.5px",color:T.text,outline:"none",fontFamily:T.font,cursor:"pointer"}}/>
            {temFiltro && (
              <button onClick={()=>{setFDataDe("");setFDataAte("");}}
                style={{border:"none",background:"none",cursor:"pointer",color:"#F24E29",fontSize:"13px",padding:"0 2px",fontWeight:700}}>✕</button>
            )}
          </div>

          {/* Ordenação */}
          <div style={{display:"flex",gap:"4px"}}>
            {[["recente","🕐 Recente"],["maior","↑ Maior valor"],["menor","↓ Menor valor"]].map(([k,l])=>(
              <button key={k} onClick={()=>setOrdenar(k)}
                style={{padding:"6px 12px",borderRadius:"20px",fontSize:"12px",fontWeight:600,cursor:"pointer",border:"none",
                  background:ordenar===k?T.primary:"transparent",
                  color:ordenar===k?"#fff":T.textSub,
                  outline:ordenar!==k?`1.5px solid ${T.border}`:"none",
                  transition:"all .12s"}}>
                {l}
              </button>
            ))}
          </div>
        </div>
      </div>

      {/* Cards resumo */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:"14px",marginBottom:"20px"}}>
        {[
          {label:"Total Emitido",  value:fmt(totalGeral),    icon:"💰",c:"#1A5173",bg:"#C4DDF2"},
          {label:"A Pagar",        value:fmt(totalPendente), icon:"⏳",c:"#d97706",bg:"#fffbeb"},
          {label:"Total Pago",     value:fmt(totalPago),     icon:"✅",c:"#059669",bg:"#ecfdf5"},
          {label:"Total Parcelas", value:totalParcelas,      icon:"📄",c:"#7c3aed",bg:"#f5f3ff"},
        ].map(card=>(
          <div key={card.label} style={{background:T.surface,borderRadius:T.radius,padding:"16px 18px",boxShadow:"0 1px 3px rgba(0,0,0,.05)",border:`1px solid ${T.border}`}}>
            <div style={{display:"flex",alignItems:"center",gap:"8px",marginBottom:"8px"}}>
              <div style={{width:"30px",height:"30px",borderRadius:"7px",background:card.bg,display:"flex",alignItems:"center",justifyContent:"center",fontSize:"15px"}}>{card.icon}</div>
              <span style={{fontSize:"12px",color:T.textSub,fontWeight:500}}>{card.label}</span>
            </div>
            <div style={{fontSize:"20px",fontWeight:800,color:T.text}}>{card.value}</div>
          </div>
        ))}
      </div>

      {/* Abas */}
      <div style={{display:"flex",gap:"4px",marginBottom:"16px",background:T.surface,borderRadius:"10px",padding:"4px",boxShadow:"0 1px 3px rgba(0,0,0,.05)",border:`1px solid ${T.border}`,width:"fit-content"}}>
        {[["notas","📄 Notas Fiscais"],["resumo","📊 Resumo"],["vencimentos","📅 Próx. Vencimentos"]].map(([k,l])=>(
          <button key={k} onClick={()=>setAba(k)} style={{padding:"7px 18px",borderRadius:"7px",border:"none",cursor:"pointer",fontSize:"13px",fontWeight:600,transition:"all .15s",background:aba===k?"#1A5173":"transparent",color:aba===k?"#fff":"#475569"}}>{l}</button>
        ))}
      </div>

      {/* ── ABA NOTAS FISCAIS ── */}
      {aba==="notas" && (
        <div style={{background:T.surface,borderRadius:T.radius,boxShadow:"0 1px 4px rgba(0,0,0,.06)",border:`1px solid ${T.border}`,overflow:"hidden"}}>
          {notasOrdenadas.length===0 ? (
            <div style={{padding:"48px",textAlign:"center",color:T.textMuted}}>
              {temFiltro ? "Nenhuma nota no período selecionado" : "Nenhuma nota fiscal encontrada para este fornecedor"}
            </div>
          ) : (
            <>
              <div style={{padding:"12px 16px",background:T.bg,borderBottom:`1px solid ${T.border}`,fontSize:"12.5px",color:T.textMuted}}>
                {notasOrdenadas.length} nota(s) · Total: <strong style={{color:T.text}}>{fmt(totalGeral)}</strong>
                {temFiltro && <span style={{marginLeft:"8px",color:"#F24E29",fontWeight:600}}>· Filtro ativo</span>}
              </div>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:"13.5px"}}>
                <thead>
                  <tr style={{background:T.bg}}>
                    <th style={{padding:"11px 14px",width:"32px",borderBottom:"1px solid #e2e8f0"}}></th>
                    {["Número NF","Empresa","Emissão","Valor Total","Parcelas","Status"].map(h=>(
                      <th key={h} style={{padding:"11px 16px",textAlign:"left",fontWeight:700,color:T.textSub,fontSize:"11.5px",textTransform:"uppercase",letterSpacing:".05em",borderBottom:"1px solid #e2e8f0",whiteSpace:"nowrap"}}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {notasOrdenadas.map((n,i)=>{
                    const st  = statusNota(n);
                    const exp = expandedId===n.id;
                    return (
                      <React.Fragment key={n.id}>
                        <tr onClick={()=>setExpandedId(exp?null:n.id)}
                          style={{borderBottom:exp?"none":i<notasOrdenadas.length-1?"1px solid #f1f5f9":"none",cursor:"pointer",
                            background: lastAddedId===n.id ? "#dbeafe" : exp ? "#f0f6ff" : "",
                            transition:"background 1.5s ease",
                          }}
                          onMouseEnter={e=>{if(!exp&&lastAddedId!==n.id)e.currentTarget.style.background="#fafbff"}}
                          onMouseLeave={e=>{if(!exp&&lastAddedId!==n.id)e.currentTarget.style.background=""}}>
                          <td style={{padding:"12px 14px",textAlign:"center"}}>
                            <span style={{fontSize:"11px",color:T.textMuted,display:"inline-block",transform:exp?"rotate(90deg)":"rotate(0deg)",transition:"transform .2s"}}>▶</span>
                          </td>
                          <td style={{padding:"12px 16px",fontWeight:700,color:T.text,fontFamily:"monospace"}}>{n.numero}</td>
                          <td style={{padding:"12px 16px"}}><span style={{background:"#f1f5f9",color:"#334155",borderRadius:"6px",padding:"2px 8px",fontSize:"12px",fontWeight:600}}>{n.empresa}</span></td>
                          <td style={{padding:"12px 16px",color:"#475569"}}>{fdate(n.emissao)}</td>
                          <td style={{padding:"12px 16px",fontWeight:700,color:T.text}}>{fmt(n.valor)}</td>
                          <td style={{padding:"12px 16px",textAlign:"center",color:"#475569"}}>{n.parcelas}x</td>
                          <td style={{padding:"12px 16px"}}><span style={{fontSize:"12px",fontWeight:600,padding:"3px 10px",borderRadius:"20px",background:st.bg,color:st.color}}>{st.label}</span></td>
                        </tr>
                        {exp && n.vencimentos.map((v,pi)=>{
                          const pago = parcelaStatusEfetivo(n,v)==="pago";
                          return (
                            <tr key={pi} style={{background:"#f8faff",borderBottom:pi===n.vencimentos.length-1?"1px solid #e2e8f0":"1px solid #eef2ff"}}>
                              <td/>
                              <td style={{padding:"9px 16px"}}><span style={{fontFamily:"monospace",fontSize:"12px",color:T.textMuted}}>{n.numero} <strong>P{pi+1}</strong></span></td>
                              <td colSpan={2} style={{padding:"9px 16px",fontSize:"12.5px",color:"#475569"}}>{fdate(v)}</td>
                              <td style={{padding:"9px 16px",fontWeight:700,color:pago?"#059669":"#374151"}}>{fmt(n.valor/n.parcelas)}</td>
                              <td colSpan={2} style={{padding:"9px 16px"}}>
                                {pago
                                  ? <span style={{fontSize:"12px",fontWeight:600,color:"#059669",background:"#ecfdf5",padding:"2px 8px",borderRadius:"20px"}}>✓ Pago</span>
                                  : <span style={{fontSize:"12px",fontWeight:600,color:v<td?"#F24E29":v===td?"#d97706":"#94a3b8"}}>
                                      {v<td?"⚠️ Vencido":v===td?"⚠️ Vence hoje":`Vence em ${fdate(v)}`}
                                    </span>
                                }
                              </td>
                            </tr>
                          );
                        })}
                      </React.Fragment>
                    );
                  })}
                </tbody>
                <tfoot>
                  <tr style={{background:T.bg,borderTop:"2px solid #e2e8f0"}}>
                    <td colSpan={4} style={{padding:"12px 16px",fontWeight:700,color:"#374151",fontSize:"13px"}}>Total</td>
                    <td style={{padding:"12px 16px",fontWeight:800,color:T.text,fontSize:"14px"}}>{fmt(totalGeral)}</td>
                    <td colSpan={2}/>
                  </tr>
                </tfoot>
              </table>
            </>
          )}
        </div>
      )}

      {/* ── ABA RESUMO ── */}
      {aba==="resumo" && (
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"16px"}}>
          <div style={{background:T.surface,borderRadius:T.radius,padding:"22px 24px",boxShadow:"0 1px 4px rgba(0,0,0,.06)",border:`1px solid ${T.border}`}}>
            <p style={{margin:"0 0 4px",fontSize:"14px",fontWeight:700,color:T.text}}>Distribuição por Empresa</p>
            <p style={{margin:"0 0 20px",fontSize:"12px",color:T.textMuted}}>Volume de compras por destino</p>
            {porEmpresa.every(e=>e.total===0) ? (
              <p style={{color:T.textMuted,fontSize:"13px",textAlign:"center",padding:"20px 0"}}>Nenhuma NF no período</p>
            ) : (
              <div style={{display:"flex",flexDirection:"column",gap:"16px"}}>
                {porEmpresa.map(e=>(
                  <div key={e.name}>
                    <div style={{display:"flex",justifyContent:"space-between",marginBottom:"5px"}}>
                      <span style={{fontSize:"13px",fontWeight:600,color:"#374151"}}>{e.name}</span>
                      <div>
                        <span style={{fontSize:"13px",fontWeight:700,color:T.text}}>{fmt(e.total)}</span>
                        <span style={{marginLeft:"8px",fontSize:"11px",color:T.textMuted}}>{e.count} NF(s)</span>
                      </div>
                    </div>
                    <div style={{height:"8px",borderRadius:"99px",background:"#f1f5f9",overflow:"hidden"}}>
                      <div style={{height:"100%",width:`${(e.total/maxEmp)*100}%`,borderRadius:"99px",background:e.color,transition:"width .4s"}}/>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
          <div style={{background:T.surface,borderRadius:T.radius,padding:"22px 24px",boxShadow:"0 1px 4px rgba(0,0,0,.06)",border:`1px solid ${T.border}`}}>
            <p style={{margin:"0 0 4px",fontSize:"14px",fontWeight:700,color:T.text}}>Status das Notas</p>
            <p style={{margin:"0 0 20px",fontSize:"12px",color:T.textMuted}}>Situação atual das NFs no período</p>
            {notasFiltradas.length===0 ? (
              <p style={{color:T.textMuted,fontSize:"13px",textAlign:"center",padding:"20px 0"}}>Nenhuma NF no período</p>
            ) : (
              <div style={{display:"flex",flexDirection:"column",gap:"10px"}}>
                {[
                  {key:"aberto",    label:"Aberto",     color:"#1547a0",bg:"#dbeafe"},
                  {key:"vencido",   label:"Vencido",    color:"#92400e",bg:"#fef3c7"},
                  {key:"incompleto",label:"Incompleto", color:"#b42318",bg:"#fdecea"},
                  {key:"quitado",   label:"Quitado",    color:"#166534",bg:"#dcfce7"},
                ].map(v=>{
                  const nfs = notasFiltradas.filter(n=>statusNota(n).key===v.key);
                  if(!nfs.length) return null;
                  return (
                    <div key={v.key} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"10px 14px",borderRadius:"10px",background:v.bg}}>
                      <div style={{display:"flex",alignItems:"center",gap:"8px"}}>
                        <div style={{width:"8px",height:"8px",borderRadius:"50%",background:v.color}}/>
                        <span style={{fontSize:"13px",fontWeight:600,color:v.color}}>{v.label}</span>
                        <span style={{fontSize:"12px",color:v.color,opacity:.7}}>{nfs.length} NF(s)</span>
                      </div>
                      <span style={{fontSize:"13px",fontWeight:700,color:v.color}}>{fmt(nfs.reduce((s,n)=>s+n.valor,0))}</span>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        </div>
      )}

      {/* ── ABA PRÓXIMOS VENCIMENTOS ── */}
      {aba==="vencimentos" && (
        <div style={{background:T.surface,borderRadius:T.radius,boxShadow:"0 1px 4px rgba(0,0,0,.06)",border:`1px solid ${T.border}`,overflow:"hidden"}}>
          {proximos.length===0 ? (
            <div style={{padding:"48px",textAlign:"center",color:T.textMuted}}>Nenhum vencimento futuro encontrado</div>
          ) : (
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:"13.5px"}}>
              <thead>
                <tr style={{background:T.bg}}>
                  {["Número NF","Empresa","Parcela","Vencimento","Valor","Situação"].map(h=>(
                    <th key={h} style={{padding:"11px 18px",textAlign:"left",fontWeight:700,color:T.textSub,fontSize:"11.5px",textTransform:"uppercase",letterSpacing:".05em",borderBottom:"1px solid #e2e8f0",whiteSpace:"nowrap"}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {proximos.map((p,i)=>{
                  const dias = Math.ceil((new Date(p.venc)-new Date(td))/(1000*60*60*24));
                  const cor  = p.pago?"#059669":dias<=3?"#F24E29":dias<=7?"#d97706":"#059669";
                  return (
                    <tr key={`${p.notaId}-${p.parcela}`} style={{borderBottom:i<proximos.length-1?"1px solid #f1f5f9":"none"}}
                      onMouseEnter={e=>e.currentTarget.style.background="#fafbff"}
                      onMouseLeave={e=>e.currentTarget.style.background=""}>
                      <td style={{padding:"13px 18px",fontWeight:700,color:T.text,fontFamily:"monospace"}}>{p.numero}</td>
                      <td style={{padding:"13px 18px"}}><span style={{background:"#f1f5f9",color:"#334155",borderRadius:"6px",padding:"2px 8px",fontSize:"12px",fontWeight:600}}>{p.empresa}</span></td>
                      <td style={{padding:"13px 18px",color:T.textSub}}>{p.parcela}/{p.total}</td>
                      <td style={{padding:"13px 18px",fontWeight:600,color:T.text}}>{fdate(p.venc)}</td>
                      <td style={{padding:"13px 18px",fontWeight:700,color:T.text}}>{fmt(p.valor)}</td>
                      <td style={{padding:"13px 18px"}}>
                        {p.pago
                          ? <span style={{fontSize:"12px",fontWeight:600,color:"#059669"}}>✓ Pago</span>
                          : <span style={{fontSize:"12.5px",fontWeight:700,color:cor}}>{dias===0?"Hoje":dias===1?"Amanhã":`Em ${dias} dias`}</span>
                        }
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          )}
        </div>
      )}
    </div>
  );
}


// ═══════════════════════════════════════════════════════════════════════════
//  FORNECEDORES
// ═══════════════════════════════════════════════════════════════════════════
function FornecedorModal({ onClose, onSave, editData, initialNome="" }) {
  const [form, setForm] = useState(editData || { nome: initialNome });
  const [errors, setErrors] = useState({});
  const set = (f,v) => { setForm(p=>({...p,[f]:v})); setErrors(p=>({...p,[f]:undefined})); };

  function handleSave() {
    const e={};
    if(!form.nome.trim()) e.nome="Obrigatório";
    setErrors(e);
    if(Object.keys(e).length) return;
    onSave({...form, id: editData?.id||Date.now()});
  }

  return (
    <div style={{position:"fixed",inset:0,background:"rgba(15,23,42,0.5)",backdropFilter:"blur(4px)",zIndex:70,display:"flex",alignItems:"center",justifyContent:"center",padding:"16px"}}>
      <div style={{background:T.surface,borderRadius:"16px",width:"100%",maxWidth:"440px",boxShadow:"0 24px 64px rgba(0,0,0,0.18)",fontFamily:T.font}}>
        <div style={{padding:"22px 26px 0",display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
          <h2 style={{margin:0,fontSize:"18px",fontWeight:800,color:T.text}}>{editData?"Editar Fornecedor":"Novo Fornecedor"}</h2>
          <button onClick={onClose} style={{border:"none",background:"none",cursor:"pointer",fontSize:"22px",color:T.textMuted}}>×</button>
        </div>
        <div style={{padding:"18px 26px 26px",display:"flex",flexDirection:"column",gap:"14px"}}>
          <div>
            <label style={S.label}>Nome</label>
            <input value={form.nome} onChange={e=>set("nome",e.target.value)} placeholder="Nome do fornecedor" autoFocus style={{...S.input,borderColor:errors.nome?"#F24E29":"#e2e8f0"}}/>
            {errors.nome&&<span style={S.error}>{errors.nome}</span>}
          </div>
          <div style={{display:"flex",gap:"10px",marginTop:"4px"}}>
            <button onClick={onClose} style={{flex:1,padding:"11px",borderRadius:"9px",border:"1.5px solid #e2e8f0",background:T.bg,color:"#475569",fontWeight:600,fontSize:"13.5px",cursor:"pointer"}}>Cancelar</button>
            <button onClick={handleSave} style={{flex:1,padding:"11px",borderRadius:"9px",border:"none",background:"linear-gradient(135deg,#1A5173,#1A5173)",color:"#fff",fontWeight:700,fontSize:"13.5px",cursor:"pointer",boxShadow:"0 4px 14px rgba(37,99,235,.3)"}}>Salvar</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function FornecedoresPage({ fornecedores, setFornecedores, saveFornecedor, notas, lastAddedId }) {
  const [showModal, setShowModal] = useState(false);
  const [editForn, setEditForn]   = useState(null);
  const [deleteId, setDeleteId]   = useState(null);
  const [blockDelete, setBlockDelete] = useState(null);
  const [search, setSearch]       = useState("");
  const [relatorio, setRelatorio] = useState(null); // fornecedor selecionado para relatório

  async function handleSave(f) {
    await saveFornecedor(f);
    setShowModal(false); setEditForn(null);
  }

  // Se há um fornecedor selecionado, mostra o relatório
  if (relatorio) {
    return <RelatorioFornecedor fornecedor={relatorio} notas={notas} onVoltar={()=>setRelatorio(null)} lastAddedId={lastAddedId}/>;
  }

  const filtered = fornecedores
    .filter(f => !search || f.nome.toLowerCase().includes(search.toLowerCase()))
    .sort((a,b) => a.nome.localeCompare(b.nome, "pt-BR"));

  // Estatísticas por fornecedor
  function statsFor(forn) {
    const nfs = notas.filter(n => n.fornecedorId === forn.id);
    return {
      total: nfs.reduce((s,n)=>s+n.valor,0),
      count: nfs.length,
      pendente: nfs.filter(n=>statusNota(n).key!=="quitado").reduce((s,n)=>s+n.valor,0),
    };
  }

  return (
    <div style={{padding:"28px 36px",fontFamily:T.font}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:"24px"}}>
        <div>
          <h1 style={{margin:0,fontSize:"22px",fontWeight:800,color:T.text}}>Fornecedores</h1>
          <p style={{margin:"3px 0 0",fontSize:"13px",color:T.textSub}}>{fornecedores.length} fornecedor(es) cadastrado(s)</p>
        </div>
        <button onClick={()=>setShowModal(true)} style={{padding:"9px 18px",borderRadius:T.radiusSm,border:"none",background:"linear-gradient(135deg,#1A5173,#1A5173)",color:"#fff",fontWeight:700,fontSize:"13px",cursor:"pointer",boxShadow:"0 4px 12px rgba(37,99,235,.3)"}}>+ Novo Fornecedor</button>
      </div>

      {/* Busca */}
      <div style={{marginBottom:"16px"}}>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Buscar por nome ou código..." style={{...S.input,maxWidth:"320px",background:T.surface,boxShadow:"0 1px 3px rgba(0,0,0,.05)"}}/>
      </div>

      {/* Lista */}
      {filtered.length===0 ? (
        <div style={{background:T.surface,borderRadius:T.radius,padding:"60px",textAlign:"center",border:`1px solid ${T.border}`,boxShadow:"0 1px 4px rgba(0,0,0,.06)"}}>
          <div style={{fontSize:"40px",marginBottom:"12px"}}>🏢</div>
          <p style={{margin:0,fontSize:"15px",fontWeight:600,color:T.textMuted}}>Nenhum fornecedor cadastrado</p>
          <p style={{margin:"6px 0 16px",fontSize:"13px",color:"#cbd5e1"}}>Cadastre fornecedores para vinculá-los às notas fiscais</p>
          <button onClick={()=>setShowModal(true)} style={{padding:"9px 18px",borderRadius:T.radiusSm,border:"none",background:"#1A5173",color:"#fff",fontWeight:700,fontSize:"13px",cursor:"pointer"}}>+ Novo Fornecedor</button>
        </div>
      ) : (
        <div style={{background:T.surface,borderRadius:T.radius,boxShadow:"0 1px 4px rgba(0,0,0,.06)",border:`1px solid ${T.border}`,overflow:"hidden"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:"13.5px"}}>
            <thead>
              <tr style={{background:T.bg}}>
                {["Código","Nome","Notas Fiscais","Total Emitido","Pendente","Ações"].map(h=>(
                  <th key={h} style={{padding:"11px 18px",textAlign:"left",fontWeight:700,color:T.textSub,fontSize:"11.5px",textTransform:"uppercase",letterSpacing:".05em",borderBottom:"1px solid #e2e8f0",whiteSpace:"nowrap"}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map((f,i)=>{
                const st = statsFor(f);
                return (
                  <tr key={f.id} style={{borderBottom:i<filtered.length-1?"1px solid #f1f5f9":"none"}}
                    onMouseEnter={e=>e.currentTarget.style.background="#fafbff"}
                    onMouseLeave={e=>e.currentTarget.style.background=""}>
                    <td style={{padding:"13px 18px"}}>

                    </td>
                    <td style={{padding:"13px 18px",fontWeight:600,color:T.text}}>{f.nome}</td>
                    <td style={{padding:"13px 18px",color:"#475569",textAlign:"center"}}>
                      <span style={{background:"#C4DDF2",color:"#1A5173",borderRadius:"20px",padding:"2px 10px",fontSize:"12px",fontWeight:700}}>{st.count}</span>
                    </td>
                    <td style={{padding:"13px 18px",fontWeight:700,color:T.text}}>{fmt(st.total)}</td>
                    <td style={{padding:"13px 18px"}}>
                      {st.pendente>0
                        ? <span style={{fontSize:"13px",fontWeight:600,color:"#d97706"}}>{fmt(st.pendente)}</span>
                        : <span style={{fontSize:"13px",color:T.textMuted}}>—</span>}
                    </td>
                    <td style={{padding:"13px 18px"}}>
                      <div style={{display:"flex",gap:"5px"}}>
                        <button onClick={()=>setRelatorio(f)} style={{padding:"5px 10px",borderRadius:"6px",border:"1.5px solid #dbeafe",background:"#C4DDF2",cursor:"pointer",fontSize:"12px",fontWeight:600,color:"#1A5173"}}>📊 Relatório</button>
                        <button onClick={()=>setEditForn({...f})} style={{padding:"5px 10px",borderRadius:"6px",border:"1.5px solid #e2e8f0",background:T.bg,cursor:"pointer",fontSize:"12px",fontWeight:600,color:"#475569"}}>✏️ Editar</button>
                        <button onClick={()=>{
                          const nfsVinculadas = notas.filter(n=>n.fornecedorId===f.id).length;
                          if(nfsVinculadas>0){ setBlockDelete({nome:f.nome, count:nfsVinculadas}); }
                          else { setDeleteId(f.id); }
                        }} style={{padding:"5px 10px",borderRadius:"6px",border:"1.5px solid #fee2e2",background:"#fff5f5",cursor:"pointer",fontSize:"12px",fontWeight:600,color:"#F24E29"}}>🗑</button>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {(showModal||editForn)&&<FornecedorModal onClose={()=>{setShowModal(false);setEditForn(null);}} onSave={handleSave} editData={editForn}/>}

      {deleteId&&(
        <div style={{position:"fixed",inset:0,background:"rgba(15,23,42,.5)",backdropFilter:"blur(4px)",zIndex:60,display:"flex",alignItems:"center",justifyContent:"center"}}>
          <div style={{background:T.surface,borderRadius:T.radius,padding:"28px",maxWidth:"340px",width:"90%",textAlign:"center",boxShadow:"0 24px 60px rgba(0,0,0,.18)"}}>
            <div style={{fontSize:"34px",marginBottom:"10px"}}>🗑️</div>
            <h3 style={{margin:"0 0 8px",color:T.text}}>Excluir fornecedor?</h3>
            <p style={{color:T.textSub,fontSize:"13.5px",marginBottom:"20px"}}>Esta ação não pode ser desfeita.</p>
            <div style={{display:"flex",gap:"10px"}}>
              <button onClick={()=>setDeleteId(null)} style={{flex:1,padding:"10px",borderRadius:T.radiusSm,border:"1.5px solid #e2e8f0",background:T.bg,color:"#475569",fontWeight:600,cursor:"pointer"}}>Cancelar</button>
              <button onClick={()=>{setFornecedores(fs=>fs.filter(f=>f.id!==deleteId));setDeleteId(null);}} style={{flex:1,padding:"10px",borderRadius:T.radiusSm,border:"none",background:"#F24E29",color:"#fff",fontWeight:700,cursor:"pointer"}}>Excluir</button>
            </div>
          </div>
        </div>
      )}

      {blockDelete&&(
        <div style={{position:"fixed",inset:0,background:"rgba(15,23,42,.5)",backdropFilter:"blur(4px)",zIndex:60,display:"flex",alignItems:"center",justifyContent:"center"}}>
          <div style={{background:T.surface,borderRadius:T.radius,padding:"28px",maxWidth:"380px",width:"90%",textAlign:"center",boxShadow:"0 24px 60px rgba(0,0,0,.18)"}}>
            <div style={{width:"52px",height:"52px",borderRadius:"50%",background:"#fef3c7",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"24px",margin:"0 auto 14px"}}>⚠️</div>
            <h3 style={{margin:"0 0 8px",color:T.text,fontSize:"16px"}}>Exclusão bloqueada</h3>
            <p style={{color:T.textSub,fontSize:"13.5px",marginBottom:"6px"}}>
              <strong style={{color:T.text}}>{blockDelete.nome}</strong> possui
            </p>
            <div style={{background:"#fef3c7",borderRadius:"10px",padding:"12px 16px",marginBottom:"20px",border:"1px solid #fde68a"}}>
              <span style={{fontSize:"22px",fontWeight:800,color:"#92400e"}}>{blockDelete.count}</span>
              <span style={{fontSize:"13px",color:"#92400e",marginLeft:"6px",fontWeight:600}}>nota{blockDelete.count>1?"s":""} fiscal vinculada{blockDelete.count>1?"s":""}</span>
            </div>
            <p style={{color:T.textMuted,fontSize:"12.5px",marginBottom:"20px"}}>
              Exclua ou transfira as notas fiscais deste fornecedor antes de removê-lo.
            </p>
            <button onClick={()=>setBlockDelete(null)} style={{width:"100%",padding:"10px",borderRadius:T.radiusSm,border:"none",background:T.primary,color:"#fff",fontWeight:700,cursor:"pointer",fontSize:"14px"}}>Entendido</button>
          </div>
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════
//  PÁGINA DE DETALHES DE VENCIMENTOS
// ═══════════════════════════════════════════════════════════════════════════
function VencimentosDetalhe({ titulo, notas, tipo, onVoltar }) {
  const td = today(), tm = tomorrow(), ew = endOfWeek();

  // Monta parcelas individuais do período
  const parcelas = useMemo(() => {
    const list = [];
    notas.forEach(n => {
      n.vencimentos.forEach((venc, i) => {
        const incluir =
          tipo==="hoje"   ? venc===td :
          tipo==="amanha" ? venc===tm :
          tipo==="semana" ? venc>=td && venc<=ew : false;
        if (!incluir) return;
        if (parcelaStatusEfetivo(n, venc)==="pago") return;
        if (statusNota(n).key==="quitado") return;
        list.push({
          notaId: n.id, numero: n.numero, fornecedor: n.fornecedor,
          empresa: n.empresa, venc,
          parcela: i+1, totalParcelas: n.parcelas,
          valor: n.valor / n.parcelas,
          boletosRecebidos: n.boletosRecebidos,
        });
      });
    });
    return list.sort((a,b)=>a.venc.localeCompare(b.venc));
  }, [notas, tipo]);

  const totalGeral = parcelas.reduce((s,p)=>s+p.valor,0);
  const accent = tipo==="hoje"?"#F24E29":tipo==="amanha"?"#f59e0b":"#1A5173";

  return (
    <div style={{padding:"28px 36px",fontFamily:T.font}}>
      {/* Cabeçalho */}
      <div style={{display:"flex",alignItems:"center",gap:"14px",marginBottom:"24px"}}>
        <button onClick={onVoltar} style={{padding:"7px 14px",borderRadius:T.radiusSm,border:"1.5px solid #e2e8f0",background:T.bg,color:"#475569",fontWeight:600,fontSize:"13px",cursor:"pointer"}}>
          ← Voltar
        </button>
        <div style={{width:"1px",height:"28px",background:"#e2e8f0"}}/>
        <div>
          <h1 style={{margin:0,fontSize:"20px",fontWeight:800,color:T.text}}>{titulo}</h1>
          <p style={{margin:0,fontSize:"12.5px",color:T.textMuted}}>{parcelas.length} boleto(s) · Total: <strong style={{color:T.text}}>{fmt(totalGeral)}</strong></p>
        </div>
      </div>

      {/* Tabela */}
      <div style={{background:T.surface,borderRadius:T.radius,boxShadow:"0 1px 4px rgba(0,0,0,.06)",border:`1px solid ${T.border}`,overflow:"hidden"}}>
        {parcelas.length===0 ? (
          <div style={{padding:"60px",textAlign:"center",color:T.textMuted}}>
            <div style={{fontSize:"36px",marginBottom:"10px"}}>✅</div>
            <p style={{margin:0,fontWeight:600}}>Nenhum boleto vencendo neste período</p>
          </div>
        ) : (
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:"13.5px"}}>
            <thead>
              <tr style={{background:T.bg}}>
                {["Número NF","Fornecedor","Empresa","Parcela","Vencimento","Valor","Boleto"].map(h=>(
                  <th key={h} style={{padding:"11px 18px",textAlign:"left",fontWeight:700,color:T.textSub,fontSize:"11.5px",textTransform:"uppercase",letterSpacing:".05em",borderBottom:"1px solid #e2e8f0",whiteSpace:"nowrap"}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {parcelas.map((p,i)=>{
                const st = {aberto:{label:"Aberto",color:"#1547a0",bg:"#dbeafe"},incompleto:{label:"Incompleto",color:"#b42318",bg:"#fdecea"},quitado:{label:"Quitado",color:"#166534",bg:"#dcfce7"}}[p.status] || statusNota({boletosRecebidos:true,vencimentos:[],parcelasPagas:[]});
                const boletoOk = p.boletosRecebidos;
                return (
                  <tr key={`${p.notaId}-${p.parcela}`}
                    style={{borderBottom:i<parcelas.length-1?"1px solid #f1f5f9":"none",background:!boletoOk?"#fffbeb":""}}
                    onMouseEnter={e=>e.currentTarget.style.background=!boletoOk?"#fef9ee":"#fafbff"}
                    onMouseLeave={e=>e.currentTarget.style.background=!boletoOk?"#fffbeb":""}>
                    <td style={{padding:"13px 18px",fontWeight:700,color:T.text,fontFamily:"monospace"}}>{p.numero}</td>
                    <td style={{padding:"13px 18px",fontWeight:600,color:T.text}}>{fnome(p.fornecedor)}</td>
                    <td style={{padding:"13px 18px"}}>
                      <span style={{background:"#f1f5f9",color:"#334155",borderRadius:"6px",padding:"2px 8px",fontSize:"12px",fontWeight:600}}>{p.empresa}</span>
                    </td>
                    <td style={{padding:"13px 18px",color:T.textSub}}>{p.parcela}/{p.totalParcelas}</td>
                    <td style={{padding:"13px 18px",fontWeight:600,color:accent}}>{fdate(p.venc)}</td>
                    <td style={{padding:"13px 18px",fontWeight:800,color:T.text}}>{fmt(p.valor)}</td>
                    <td style={{padding:"13px 18px"}}>
                      {boletoOk
                        ? <span style={{fontSize:"12px",fontWeight:600,padding:"3px 10px",borderRadius:"20px",background:"#d1fae5",color:"#065f46"}}>✓ Recebido</span>
                        : <span style={{fontSize:"12px",fontWeight:600,padding:"3px 10px",borderRadius:"20px",background:"#fef3c7",color:"#92400e"}}>⚠ Pendente</span>}
                    </td>
                  </tr>
                );
              })}
            </tbody>
            {/* Rodapé com total */}
            <tfoot>
              <tr style={{background:T.bg,borderTop:"2px solid #e2e8f0"}}>
                <td colSpan={5} style={{padding:"12px 18px",fontWeight:700,color:"#374151",fontSize:"13px"}}>Total do período</td>
                <td style={{padding:"12px 18px",fontWeight:800,color:T.text,fontSize:"14px"}}>{fmt(totalGeral)}</td>
                <td/>
              </tr>
            </tfoot>
          </table>
        )}
      </div>
    </div>
  );
}


// ═══════════════════════════════════════════════════════════════════════════
//  LOGIN
// ═══════════════════════════════════════════════════════════════════════════
function LoginPage({ onLogin }) {
  const [email, setEmail]     = useState("lukaskowaa@gmail.com");
  const [senha, setSenha]     = useState("");
  const [erro, setErro]       = useState("");
  const [loading, setLoading] = useState(false);

  async function handleLogin() {
    setErro(""); setLoading(true);
    try {
      await sbLogin(email, senha);
      onLogin();
    } catch(e) {
      setErro("Email ou senha incorretos.");
    } finally {
      setLoading(false);
    }
  }

  function onKey(e) { if(e.key==="Enter") handleLogin(); }

  return (
    <div style={{minHeight:"100vh",background:T.bg,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:T.font}}>
      <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@400;500;600&display=swap" rel="stylesheet"/>
      <div style={{background:T.surface,borderRadius:T.radius,boxShadow:T.shadowMd,padding:"40px 44px",width:"100%",maxWidth:"400px",border:`1px solid ${T.border}`}}>
        <div style={{display:"flex",alignItems:"center",gap:"10px",marginBottom:"28px"}}>
          <div style={{width:"36px",height:"36px",borderRadius:"9px",background:T.primary,display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontWeight:800,fontSize:"18px"}}>K</div>
          <div>
            <div style={{fontSize:"16px",fontWeight:700,color:T.text}}>Kowalsky</div>
            <div style={{fontSize:"12px",color:T.textMuted}}>Gestão Financeira</div>
          </div>
        </div>
        <h2 style={{margin:"0 0 24px",fontSize:"20px",fontWeight:700,color:T.text}}>Entrar</h2>
        <div style={{marginBottom:"14px"}}>
          <label style={S.label}>Email</label>
          <input value={email} onChange={e=>setEmail(e.target.value)} onKeyDown={onKey}
            type="email" placeholder="seu@email.com"
            style={{...S.input, borderColor: erro ? T.danger : T.border}}/>
        </div>
        <div style={{marginBottom:"22px"}}>
          <label style={S.label}>Senha</label>
          <input value={senha} onChange={e=>setSenha(e.target.value)} onKeyDown={onKey}
            type="password" placeholder="••••••••"
            style={{...S.input, borderColor: erro ? T.danger : T.border}}/>
          {erro && <span style={S.error}>{erro}</span>}
        </div>
        <button onClick={handleLogin} disabled={loading}
          style={{width:"100%",padding:"11px",borderRadius:T.radiusSm,border:"none",background:loading?T.primaryLight:T.primary,color:loading?T.primary:"#fff",fontWeight:700,fontSize:"14px",cursor:loading?"not-allowed":"pointer",fontFamily:T.font,transition:"all .15s"}}>
          {loading ? "Entrando..." : "Entrar"}
        </button>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════
//  ROOT
// ═══════════════════════════════════════════════════════════════════════════
const SEED = [
  // ══════════════════════════════════════════════════════════════════════
  // KOWALSKY — histórico completo desde 2025
  // ══════════════════════════════════════════════════════════════════════

  // NF antiga totalmente paga (boleto recebido + todos vencimentos no passado → pago automático)
  {id:1,  fornecedor:"FORN001 - Brambila Distribuidora", fornecedorId:1, empresa:"Kowalsky", numero:"NF-1001", emissao:"2025-08-01", valor:12000.00, boletosRecebidos:true,  parcelas:3, vencimentos:["2025-09-01","2025-10-01","2025-11-01"], status:"pago",    observacao:"Contrato Q3 2025 - quitado"},
  // NF paga manualmente via status
  {id:2,  fornecedor:"FORN002 - Pereira Suprimentos",    fornecedorId:2, empresa:"Kowalsky", numero:"NF-1002", emissao:"2025-10-15", valor:3400.00,  boletosRecebidos:true,  parcelas:1, vencimentos:["2025-11-15"], status:"pago",    observacao:"Pago via PIX"},
  // NF rejeitada — divergência fiscal
  {id:3,  fornecedor:"FORN003 - TechFrio Equipamentos",  fornecedorId:3, empresa:"Kowalsky", numero:"NF-1003", emissao:"2025-11-20", valor:8750.00,  boletosRecebidos:false, parcelas:2, vencimentos:["2025-12-20","2026-01-20"], status:"rejeitado", observacao:"CNPJ divergente — devolvida ao fornecedor"},
  // NF com parcelas mistas: primeiras já pagas automaticamente, últimas futuras
  {id:4,  fornecedor:"FORN001 - Brambila Distribuidora", fornecedorId:1, empresa:"Kowalsky", numero:"NF-1004", emissao:"2025-12-01", valor:18000.00, boletosRecebidos:true,  parcelas:6, vencimentos:["2026-01-05","2026-02-05","2026-03-05","2026-04-05","2026-05-05","2026-06-05"], status:"aprovado", observacao:"Parcelamento semestral"},
  // NF com boleto NÃO recebido e vencimento hoje → aparece em Boletos Não Recebidos E Pagamentos de Hoje
  {id:5,  fornecedor:"FORN004 - Martins & Cia Ltda",     fornecedorId:4, empresa:"Kowalsky", numero:"NF-1005", emissao:"2026-02-20", valor:5500.00,  boletosRecebidos:false, parcelas:2, vencimentos:["2026-03-06","2026-04-06"], status:"pendente",  observacao:"Boleto físico extraviado"},
  // NF de valor alto aprovada, vence amanhã
  {id:6,  fornecedor:"FORN003 - TechFrio Equipamentos",  fornecedorId:3, empresa:"Kowalsky", numero:"NF-1006", emissao:"2026-02-10", valor:42000.00, boletosRecebidos:true,  parcelas:4, vencimentos:["2026-03-07","2026-04-07","2026-05-07","2026-06-07"], status:"aprovado", observacao:"Compressor industrial"},
  // NF à vista, vence esta semana
  {id:7,  fornecedor:"FORN005 - Globo Materiais",        fornecedorId:5, empresa:"Kowalsky", numero:"NF-1007", emissao:"2026-03-01", valor:1250.00,  boletosRecebidos:true,  parcelas:1, vencimentos:["2026-03-08"], status:"pendente",  observacao:""},
  // NF pequena valor baixo já paga
  {id:99, fornecedor:"FORN001 - Brambila Distribuidora", fornecedorId:1, empresa:"Kowalsky", numero:"NF-TESTE", emissao:"2026-01-10", valor:6000.00, boletosRecebidos:true, parcelas:3, vencimentos:["2026-01-20","2026-02-20","2026-03-20"], parcelasPagas:[], status:"aprovado", observacao:"TESTE — parcelas vencidas não pagas"},
  {id:8,  fornecedor:"FORN002 - Pereira Suprimentos",    fornecedorId:2, empresa:"Kowalsky", numero:"NF-1008", emissao:"2025-09-10", valor:480.00,   boletosRecebidos:true,  parcelas:1, vencimentos:["2025-10-10"], status:"pago",    observacao:"Material de escritório"},

  // ══════════════════════════════════════════════════════════════════════
  // EVALDO
  // ══════════════════════════════════════════════════════════════════════

  // NF com 12 parcelas (máximo) — teste do limite
  {id:9,  fornecedor:"FORN003 - TechFrio Equipamentos",  fornecedorId:3, empresa:"Evaldo",   numero:"NF-2001", emissao:"2025-07-01", valor:60000.00, boletosRecebidos:true,  parcelas:12, vencimentos:["2025-08-01","2025-09-01","2025-10-01","2025-11-01","2025-12-01","2026-01-01","2026-02-01","2026-03-01","2026-04-01","2026-05-01","2026-06-01","2026-07-01"], status:"aprovado", observacao:"Contrato anual equipamentos — 12x"},
  // NF vencida no passado com boleto recebido → deve aparecer como pago automático
  {id:10, fornecedor:"FORN001 - Brambila Distribuidora", fornecedorId:1, empresa:"Evaldo",   numero:"NF-2002", emissao:"2026-01-15", valor:7200.00,  boletosRecebidos:true,  parcelas:2, vencimentos:["2026-02-15","2026-03-15"], status:"aprovado", observacao:""},
  // NF vencida sem boleto → deve aparecer em Boletos Não Recebidos como vencido
  {id:11, fornecedor:"FORN004 - Martins & Cia Ltda",     fornecedorId:4, empresa:"Evaldo",   numero:"NF-2003", emissao:"2026-01-20", valor:9300.00,  boletosRecebidos:false, parcelas:3, vencimentos:["2026-02-20","2026-03-20","2026-04-20"], status:"pendente",  observacao:"Fornecedor não enviou boletos"},
  // NF aprovada vencendo hoje
  {id:12, fornecedor:"FORN005 - Globo Materiais",        fornecedorId:5, empresa:"Evaldo",   numero:"NF-2004", emissao:"2026-02-06", valor:3600.00,  boletosRecebidos:true,  parcelas:1, vencimentos:["2026-03-06"], status:"aprovado", observacao:""},
  // NF futura parcelada
  {id:13, fornecedor:"FORN002 - Pereira Suprimentos",    fornecedorId:2, empresa:"Evaldo",   numero:"NF-2005", emissao:"2026-03-01", valor:15800.00, boletosRecebidos:false, parcelas:4, vencimentos:["2026-04-01","2026-05-01","2026-06-01","2026-07-01"], status:"pendente",  observacao:"Aguardando boletos"},
  // NF com valor decimal quebrado
  {id:14, fornecedor:"FORN001 - Brambila Distribuidora", fornecedorId:1, empresa:"Evaldo",   numero:"NF-2006", emissao:"2025-11-05", valor:4333.33,  boletosRecebidos:true,  parcelas:3, vencimentos:["2025-12-05","2026-01-05","2026-02-05"], status:"pago",    observacao:"Valor fracionado — teste de arredondamento"},
  // NF rejeitada recente
  {id:15, fornecedor:"FORN003 - TechFrio Equipamentos",  fornecedorId:3, empresa:"Evaldo",   numero:"NF-2007", emissao:"2026-02-25", valor:22000.00, boletosRecebidos:false, parcelas:2, vencimentos:["2026-03-25","2026-04-25"], status:"rejeitado", observacao:"Produto não entregue conforme pedido"},

  // ══════════════════════════════════════════════════════════════════════
  // SUPERFRIO
  // ══════════════════════════════════════════════════════════════════════

  // NF de 2025 totalmente no passado
  {id:16, fornecedor:"FORN005 - Globo Materiais",        fornecedorId:5, empresa:"Superfrio",numero:"NF-3001", emissao:"2025-06-01", valor:5600.00,  boletosRecebidos:true,  parcelas:2, vencimentos:["2025-07-01","2025-08-01"], status:"pago",    observacao:""},
  // NF grande parcelada com boleto não recebido e parcela vencendo esta semana
  {id:17, fornecedor:"FORN003 - TechFrio Equipamentos",  fornecedorId:3, empresa:"Superfrio",numero:"NF-3002", emissao:"2026-02-01", valor:38000.00, boletosRecebidos:false, parcelas:5, vencimentos:["2026-03-08","2026-04-08","2026-05-08","2026-06-08","2026-07-08"], status:"aprovado", observacao:"Câmaras frigoríficas — boletos não recebidos"},
  // NF com 1 parcela vencendo hoje com boleto recebido
  {id:18, fornecedor:"FORN004 - Martins & Cia Ltda",     fornecedorId:4, empresa:"Superfrio",numero:"NF-3003", emissao:"2026-02-06", valor:6800.00,  boletosRecebidos:true,  parcelas:1, vencimentos:["2026-03-06"], status:"aprovado", observacao:"Vence hoje — boleto em mãos"},
  // NF aprovada parcelas futuras distantes
  {id:19, fornecedor:"FORN001 - Brambila Distribuidora", fornecedorId:1, empresa:"Superfrio",numero:"NF-3004", emissao:"2026-03-01", valor:11200.00, boletosRecebidos:true,  parcelas:4, vencimentos:["2026-05-01","2026-06-01","2026-07-01","2026-08-01"], status:"pendente",  observacao:""},
  // NF à vista já vencida com boleto recebido → pago automaticamente
  {id:20, fornecedor:"FORN002 - Pereira Suprimentos",    fornecedorId:2, empresa:"Superfrio",numero:"NF-3005", emissao:"2026-02-10", valor:2900.00,  boletosRecebidos:true,  parcelas:1, vencimentos:["2026-03-01"], status:"pendente",  observacao:"Deve aparecer como pago automático"},
  // NF com boleto não recebido vencendo amanhã
  {id:21, fornecedor:"FORN005 - Globo Materiais",        fornecedorId:5, empresa:"Superfrio",numero:"NF-3006", emissao:"2026-02-15", valor:4100.00,  boletosRecebidos:false, parcelas:1, vencimentos:["2026-03-07"], status:"pendente",  observacao:"URGENTE — vence amanhã sem boleto"},
  // NF de valor altíssimo
  {id:22, fornecedor:"FORN003 - TechFrio Equipamentos",  fornecedorId:3, empresa:"Superfrio",numero:"NF-3007", emissao:"2025-12-01", valor:95000.00, boletosRecebidos:true,  parcelas:10, vencimentos:["2026-01-15","2026-02-15","2026-03-15","2026-04-15","2026-05-15","2026-06-15","2026-07-15","2026-08-15","2026-09-15","2026-10-15"], status:"aprovado", observacao:"Obra de expansão — parcelamento 10x"},
  // NF simples vencendo próxima semana
  {id:23, fornecedor:"FORN004 - Martins & Cia Ltda",     fornecedorId:4, empresa:"Superfrio",numero:"NF-3008", emissao:"2026-02-20", valor:780.00,   boletosRecebidos:true,  parcelas:1, vencimentos:["2026-03-12"], status:"pendente",  observacao:""},
  // NF com observação longa — teste de layout
  {id:24, fornecedor:"FORN002 - Pereira Suprimentos",    fornecedorId:2, empresa:"Kowalsky", numero:"NF-1009", emissao:"2026-01-10", valor:6600.00,  boletosRecebidos:false, parcelas:2, vencimentos:["2026-02-10","2026-03-10"], status:"pendente",  observacao:"Nota emitida com dados corretos porém boleto enviado para endereço errado — reenvio solicitado em 15/02"},
  // NF aprovada vencida sem boleto → deve aparecer como vencida em Boletos Não Recebidos
  {id:25, fornecedor:"FORN001 - Brambila Distribuidora", fornecedorId:1, empresa:"Evaldo",   numero:"NF-2008", emissao:"2026-01-25", valor:13500.00, boletosRecebidos:false, parcelas:3, vencimentos:["2026-02-25","2026-03-25","2026-04-25"], status:"aprovado", observacao:"3 parcelas — 1ª já vencida sem boleto"},
];


const SEED_FORNECEDORES = [
  { id:1, codigo:"FORN001", nome:"Brambila Distribuidora" },
  { id:2, codigo:"FORN002", nome:"Pereira Suprimentos"    },
  { id:3, codigo:"FORN003", nome:"TechFrio Equipamentos"  },
  { id:4, codigo:"FORN004", nome:"Martins & Cia Ltda"     },
  { id:5, codigo:"FORN005", nome:"Globo Materiais"        },
];

export default function App() {
  const [session, setSession]              = useState(getSession);
  const [loading, setLoading]              = useState(true);
  const [page,setPage]                     = useState("home");
  const [notasRaw,setNotasRaw]             = useState([]);
  const [showModal,setShowModal]           = useState(false);
  const [tarefasRaw,setTarefasRaw]         = useState([]);
  const [fornecedoresRaw,setFornecedoresRaw] = useState([]);
  const [novoFornModal,setNovoFornModal]   = useState(null);
  const [vencDetalhe,setVencDetalhe]       = useState(null);
  const [lastUpdated, setLastUpdated]      = useState(null);
  const [lastAddedId, setLastAddedId]      = useState(null);

  // ── Carregar dados do Supabase ──────────────────────────────────────────
  async function loadData(showLoader=true) {
    if (!session) return;
    if (showLoader) setLoading(true);
    try {
      const [nfs, forns, tars] = await Promise.all([
        sbFetch("/rest/v1/notas_fiscais?select=*&order=id"),
        sbFetch("/rest/v1/fornecedores?select=*&order=id"),
        sbFetch("/rest/v1/tarefas?select=*&order=id"),
      ]);
      setNotasRaw(nfs.map(dbToNota));
      setFornecedoresRaw(forns.map(dbToFornecedor));
      setTarefasRaw(tars.map(dbToTarefa));
      setLastUpdated(new Date());
    } catch(e) {
      if (e.message === "Sessão expirada") return; // já tratado pelo sbFetch
      console.error(e);
    }
    finally { if (showLoader) setLoading(false); }
  }

  useEffect(() => {
    function onExpired() { setSession(null); setNotasRaw([]); setFornecedoresRaw([]); setTarefasRaw([]); }
    window.addEventListener("sb_session_expired", onExpired);
    return () => window.removeEventListener("sb_session_expired", onExpired);
  }, []);

  useEffect(() => {
    if (!session) { setLoading(false); return; }
    loadData(true);
    // Auto-refresh a cada 60 segundos
    const interval = setInterval(() => loadData(false), 180000);
    return () => clearInterval(interval);
  }, [session]);

  // ── Setters que sincronizam com Supabase ───────────────────────────────
  async function setNotas(updater) {
    setNotasRaw(prev => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      return next;
    });
  }

  async function saveNota(nota) {
    const body = notaToDB(nota);
    // IDs from DB are small integers; Date.now() temp IDs are 13+ digits
    const isExisting = nota.id && typeof nota.id === "number" && nota.id < 1000000000;
    if (isExisting) {
      // update
      await sbFetch(`/rest/v1/notas_fiscais?id=eq.${nota.id}`, {
        method: "PATCH", body: JSON.stringify(body),
      });
    } else {
      // insert
      const [created] = await sbFetch("/rest/v1/notas_fiscais", {
        method: "POST", body: JSON.stringify(body),
      });
      return created ? dbToNota(created) : nota;
    }
    return nota;
  }

  async function deleteNota(id) {
    await sbFetch(`/rest/v1/notas_fiscais?id=eq.${id}`, { method: "DELETE" });
    setNotasRaw(prev => prev.filter(n => n.id !== id));
  }

  async function updateNotaField(id, fields) {
    const body = {};
    if ("boletosRecebidos" in fields) body.boletos_recebidos = fields.boletosRecebidos;
    if ("parcelasPagas"    in fields) body.parcelas_pagas    = fields.parcelasPagas;
    await sbFetch(`/rest/v1/notas_fiscais?id=eq.${id}`, {
      method: "PATCH", body: JSON.stringify(body),
    });
    setNotasRaw(prev => prev.map(n => n.id === id ? { ...n, ...fields } : n));
  }

  async function saveFornecedor(forn) {
    const isExisting = forn.id && typeof forn.id === 'number' && forn.id < 1000000000;
    if (isExisting) {
      await sbFetch(`/rest/v1/fornecedores?id=eq.${forn.id}`, {
        method: "PATCH", body: JSON.stringify({ nome: forn.nome }),
      });
      setFornecedoresRaw(prev => prev.map(f => f.id === forn.id ? forn : f));
    } else {
      const [created] = await sbFetch("/rest/v1/fornecedores", {
        method: "POST", body: JSON.stringify({ nome: forn.nome }),
      });
      const novo = dbToFornecedor(created);
      setFornecedoresRaw(prev => [...prev, novo]);
      return novo;
    }
    return forn;
  }

  async function saveTarefa(tarefa) {
    const body = tarefaToDB(tarefa);
    const isExisting = tarefa.id && typeof tarefa.id === "number" && tarefa.id < 1000000000;
    if (isExisting) {
      await sbFetch(`/rest/v1/tarefas?id=eq.${tarefa.id}`, {
        method: "PATCH", body: JSON.stringify(body),
      });
      setTarefasRaw(prev => prev.map(t => t.id === tarefa.id ? tarefa : t));
    } else {
      const [created] = await sbFetch("/rest/v1/tarefas", {
        method: "POST", body: JSON.stringify(body),
      });
      setTarefasRaw(prev => [...prev, dbToTarefa(created)]);
    }
  }

  async function deleteTarefa(id) {
    await sbFetch(`/rest/v1/tarefas?id=eq.${id}`, { method: "DELETE" });
    setTarefasRaw(prev => prev.filter(t => t.id !== id));
  }

  // ── Wrappers para compatibilidade com os componentes existentes ─────────
  async function setFornecedores(updater) {
    const prev = fornecedoresRaw;
    const next = typeof updater === "function" ? updater(prev) : updater;
    // INSERT — novo fornecedor
    const inseridos = next.filter(f => !prev.find(o => o.id === f.id));
    for (const f of inseridos) {
      try {
        const [created] = await sbFetch("/rest/v1/fornecedores", {
          method: "POST", body: JSON.stringify({ nome: f.nome }),
        });
        const novo = dbToFornecedor(created);
        setFornecedoresRaw(cur => [...cur, novo]);
      } catch(e) { console.error("Erro ao salvar fornecedor:", e); }
      return;
    }
    // PATCH — fornecedores alterados
    for (const f of next) {
      const old = prev.find(o => o.id === f.id);
      if (old && JSON.stringify(old) !== JSON.stringify(f)) {
        try {
          await sbFetch(`/rest/v1/fornecedores?id=eq.${f.id}`, {
            method: "PATCH", body: JSON.stringify({ nome: f.nome }),
          });
          if (old.nome !== f.nome) {
            await sbFetch(`/rest/v1/notas_fiscais?fornecedor_id=eq.${f.id}`, {
              method: "PATCH", body: JSON.stringify({ fornecedor: f.nome }),
            });
            setNotasRaw(cur => cur.map(n => n.fornecedorId === f.id ? {...n, fornecedor: f.nome} : n));
          }
        } catch(e) { console.error("Erro ao atualizar fornecedor:", e); }
      }
    }
    // DELETE — fornecedores removidos
    for (const f of prev) {
      if (!next.find(n => n.id === f.id)) {
        try {
          await sbFetch(`/rest/v1/fornecedores?id=eq.${f.id}`, { method: "DELETE" });
        } catch(e) { console.error("Erro ao deletar fornecedor:", e); }
      }
    }
    setFornecedoresRaw(next);
  }
  async function setTarefas(updater) {
    const prev = tarefasRaw;
    const next = typeof updater === "function" ? updater(prev) : updater;
    // Novos itens
    for (const t of next) {
      const old = prev.find(o => o.id === t.id);
      if (!old) {
        // insert — saveTarefa retorna o item com ID real do banco
        const body = tarefaToDB(t);
        try {
          const [created] = await sbFetch("/rest/v1/tarefas", { method:"POST", body:JSON.stringify(body) });
          // substitui o item temporário pelo item com ID real
          setTarefasRaw(cur => [...cur.filter(x=>x.id!==t.id), dbToTarefa(created)]);
        } catch(e) { console.error("Erro ao salvar tarefa:", e); }
        return;
      }
      if (JSON.stringify(old) !== JSON.stringify(t)) {
        // update
        saveTarefa(t);
      }
    }
    // Deletados
    prev.forEach(t => { if (!next.find(n => n.id === t.id)) deleteTarefa(t.id); });
    setTarefasRaw(next);
  }

  // ── Logout ───────────────────────────────────────────────────────────────
  async function handleLogout() {
    await sbLogout();
    setSession(null);
    setNotasRaw([]); setFornecedoresRaw([]); setTarefasRaw([]);
  }

  function navTo(p) { setPage(p); setShowModal(false); setVencDetalhe(null); }
  const pendentesHoje = tarefasRaw.filter(t=>tarefaAtivaHoje(t)&&t.concluidoEm!==today()).length;

  const [novoFornCallback, setNovoFornCallback] = useState(null);
  async function handleNovoFornecedor(nomeInicial, onCriado) {
    setNovoFornModal(nomeInicial||"");
    if (onCriado) setNovoFornCallback(()=>onCriado);
  }
  async function handleSaveNovoForn(f) {
    const novo = await saveFornecedor(f);
    setNovoFornModal(null);
    if (novoFornCallback) { novoFornCallback(novo); setNovoFornCallback(null); }
    return novo;
  }

  // ── Wrapper setNotas que persiste ────────────────────────────────────────
  async function setNotasPersist(updater) {
    const prev = notasRaw;
    const next = typeof updater === "function" ? updater(prev) : updater;
    // Inserções
    for (const n of next) {
      const old = prev.find(o => o.id === n.id);
      if (!old) {
        try {
          const created = await saveNota(n);
          if (created && created.id !== n.id)
            setNotasRaw(cur => [...cur.filter(x => x.id !== n.id), created]);
          else
            setNotasRaw(cur => [...cur.filter(x => x.id !== n.id), n]);
        } catch(e) { console.error("Erro ao salvar nota:", e); }
        return;
      }
      if (JSON.stringify(old) !== JSON.stringify(n)) {
        try { await saveNota(n); } catch(e) { console.error("Erro ao atualizar nota:", e); }
      }
    }
    // Exclusões
    for (const n of prev) {
      if (!next.find(x => x.id === n.id)) {
        try { await deleteNota(n.id); } catch(e) { console.error("Erro ao deletar nota:", e); }
      }
    }
    setNotasRaw(next);
  }


  if (!session) return <LoginPage onLogin={() => setSession(getSession())}/>;

  if (loading) return (
    <div style={{minHeight:"100vh",background:T.bg,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:T.font}}>
      <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@400;500;600&display=swap" rel="stylesheet"/>
      <div style={{textAlign:"center",color:T.textMuted}}>
        <div style={{fontSize:"32px",marginBottom:"12px"}}>⏳</div>
        <div style={{fontSize:"14px"}}>Carregando dados...</div>
      </div>
    </div>
  );

  return (
    <div style={{display:"flex",height:"100vh",overflow:"hidden",background:T.bg,fontFamily:T.font}}>
      <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@400;500;600&display=swap" rel="stylesheet"/>
      <Sidebar page={page} setPage={navTo} badge={pendentesHoje} onLogout={handleLogout} userEmail={session?.user?.email} lastUpdated={lastUpdated}/>
      <main style={{flex:1,overflowY:"auto"}}>
        {page==="home" && !vencDetalhe && (
          <HomePage notas={notasRaw} tarefas={tarefasRaw} setTarefas={setTarefas}
            onVerDetalhes={(tipo,titulo)=>setVencDetalhe({tipo,titulo})}/>
        )}
        {page==="home" && vencDetalhe && (
          <VencimentosDetalhe tipo={vencDetalhe.tipo} titulo={vencDetalhe.titulo} notas={notasRaw} onVoltar={()=>setVencDetalhe(null)}/>
        )}
        {page==="notas"        && <NotasFiscaisPage notas={notasRaw} setNotas={setNotasPersist} showModal={showModal} setShowModal={setShowModal} fornecedores={fornecedoresRaw} onNovoFornecedor={handleNovoFornecedor} lastAddedId={lastAddedId} setLastAddedId={setLastAddedId}/>}
        {page==="boletos"      && <BoletosNaoRecebidosPage notas={notasRaw} setNotas={setNotasPersist}/>}
        {page==="fornecedores" && <FornecedoresPage fornecedores={fornecedoresRaw} setFornecedores={setFornecedores} saveFornecedor={saveFornecedor} notas={notasRaw} lastAddedId={lastAddedId}/>}
        {page==="avisos"       && <AvisosPage tarefas={tarefasRaw} setTarefas={setTarefas}/>}
      </main>
      {page==="home"&&!vencDetalhe&&(
        <button onClick={()=>{setPage("notas");setTimeout(()=>setShowModal(true),50);}} style={{position:"fixed",bottom:"28px",right:"28px",width:"52px",height:"52px",borderRadius:"50%",background:"linear-gradient(135deg,#1A5173,#1A5173)",border:"none",color:"#fff",fontSize:"26px",cursor:"pointer",boxShadow:"0 6px 20px rgba(37,99,235,.4)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:40}}>+</button>
      )}
      {novoFornModal!==null && (
        <FornecedorModal onClose={()=>setNovoFornModal(null)} onSave={handleSaveNovoForn} initialNome={novoFornModal}/>
      )}
    </div>
  );
}
