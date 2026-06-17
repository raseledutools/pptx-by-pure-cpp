// pptx_viewer.cpp  —  PPTX Viewer + Text Editor using WebView2
// Architecture:
//   C++ side  : opens file dialog, unzips .pptx (ZIP), parses slide XML, resolves theme
//               colors + embedded images, renders to HTML.
//   WebView2  : renders slides pixel-perfect with HTML/CSS (colors, fonts, shapes, images)
//   Edit      : text runs are contenteditable; JS posts message back on every edit;
//               Ctrl+S asks C++ to patch those runs (Basic listener added).
//
// Compile (MSVC x64 — requires WebView2 NuGet package):
//   cl /std:c++17 /O2 /EHsc /utf-8 main.cpp ^
//      ole32.lib oleaut32.lib shlwapi.lib shell32.lib ^
//      WebView2LoaderStatic.lib ^
//      /Fe:pptx_viewer.exe /link /SUBSYSTEM:WINDOWS /ENTRY:wWinMainCRTStartup

#define _USE_MATH_DEFINES
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <windowsx.h>
#include <commdlg.h> // FIXED: Added missing header for GetOpenFileNameW
#include <shlwapi.h>
#include <shobjidl.h>
#include <shellapi.h>
#include <wrl.h>
#include <wrl/event.h>
#include "WebView2.h"          

#include <string>
#include <vector>
#include <map>
#include <unordered_map>
#include <fstream>
#include <sstream>
#include <algorithm>
#include <functional>
#include <cstdint>
#include <cstring>
#include <cassert>
#include <ctime>

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "shell32.lib")

using namespace Microsoft::WRL;
using namespace std;

// ─────────────────────────────────────────────────────────────────────────────
// MINI ZIP READER  (no external lib — reads .pptx which is a ZIP file)
// ─────────────────────────────────────────────────────────────────────────────
struct ZipEntry { string name; uint32_t offset; uint32_t compSize; uint32_t uncompSize; uint16_t method; };

static vector<uint8_t> readFile(const wstring& path) {
    HANDLE h = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, 0, nullptr);
    if (h == INVALID_HANDLE_VALUE) return {};
    LARGE_INTEGER sz; GetFileSizeEx(h, &sz);
    vector<uint8_t> buf(sz.QuadPart);
    DWORD rd; ReadFile(h, buf.data(), (DWORD)buf.size(), &rd, nullptr);
    CloseHandle(h);
    return buf;
}

static uint32_t ru32(const uint8_t* p){ return p[0]|(p[1]<<8)|(p[2]<<16)|(p[3]<<24); }
static uint16_t ru16(const uint8_t* p){ return p[0]|(p[1]<<8); }

// ── miniz inflate (public domain, trimmed) ───────────────────────────────────
#include <cstddef>
namespace tinfl {
struct State {
    const uint8_t* src; size_t srcLen; size_t srcPos;
    uint8_t* dst; size_t dstCap; size_t dstPos;
    uint32_t bits; int nBits;
    bool ok;
    uint32_t get(int n){
        while(nBits<n){ if(srcPos>=srcLen){ok=false;return 0;} bits|=(uint32_t)src[srcPos++]<<nBits; nBits+=8; }
        uint32_t v=bits&((1u<<n)-1); bits>>=n; nBits-=n; return v;
    }
    void put(uint8_t b){ if(dstPos<dstCap) dst[dstPos++]=b; else ok=false; }
    uint8_t back(size_t d){ return (dstPos>=d)?dst[dstPos-d]:0; }
};
static const uint8_t clcl[]={16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15};
struct HTree{ uint16_t sym[288]; uint8_t len[288]; int n;
    uint32_t decode(State& s){
        int maxLen=0; for(int i=0;i<n;i++) if(len[i]) maxLen=max(maxLen,(int)len[i]);
        uint32_t code=0; int nb=0;
        while(nb<=maxLen){
            code=(code<<1)|s.get(1); nb++;
            for(int i=0;i<n;i++) if(len[i]==nb){
                int cnt=0; uint32_t c=0;
                for(int j=0;j<i;j++) if(len[j]==nb) cnt++;
                uint32_t fc=0; int prev=0;
                for(int b=1;b<nb;b++){
                    int nc=0; for(int j=0;j<n;j++) if(len[j]==b) nc++;
                    fc=(fc+prev)<<1; prev=nc;
                }
                if(code==fc+cnt) return sym[i];
            }
        }
        s.ok=false; return 0;
    }
};
static void buildFixed(HTree& lit, HTree& dst){
    lit.n=288; for(int i=0;i<288;i++){lit.sym[i]=(uint16_t)i; lit.len[i]=(i<144)?8:(i<256)?9:(i<280)?7:8;}
    dst.n=32; for(int i=0;i<32;i++){dst.sym[i]=(uint16_t)i; dst.len[i]=5;}
}
static int inflate(const uint8_t* src,size_t srcLen,uint8_t* dst,size_t dstCap){
    State s{src,srcLen,0,dst,dstCap,0,0,0,true};
    static const int lbase[]={3,4,5,6,7,8,9,10,11,13,15,17,19,23,27,31,35,43,51,59,67,83,99,115,131,163,195,227,258};
    static const int lext[]= {0,0,0,0,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,5,5,5,5,0};
    static const int dbase[]={1,2,3,4,5,7,9,13,17,25,33,49,65,97,129,193,257,385,513,769,1025,1537,2049,3073,4097,6145,8193,12289,16385,24577};
    static const int dext[]= {0,0,0,0,1,1,2,2,3,3,4,4,5,5,6,6,7,7,8,8,9,9,10,10,11,11,12,12,13,13};
    int bfinal=0;
    while(!bfinal&&s.ok){
        bfinal=s.get(1); int btype=s.get(2);
        if(btype==0){ 
            s.bits=0;s.nBits=0;
            uint16_t len=ru16(src+s.srcPos); s.srcPos+=4;
            for(int i=0;i<len&&s.ok;i++) s.put(src[s.srcPos++]);
        } else {
            HTree lit,dist; lit.n=dist.n=0;
            if(btype==1){ buildFixed(lit,dist); }
            else{
                int hlit=s.get(5)+257, hdist=s.get(5)+1, hclen=s.get(4)+4;
                uint8_t cl[19]={};
                for(int i=0;i<hclen;i++) cl[clcl[i]]=s.get(3);
                HTree cc; cc.n=19; for(int i=0;i<19;i++){cc.sym[i]=(uint16_t)i;cc.len[i]=cl[i];}
                uint8_t lens[320]={}; int total=hlit+hdist,idx=0;
                while(idx<total&&s.ok){
                    uint32_t c=cc.decode(s);
                    if(c<16){lens[idx++]=(uint8_t)c;}
                    else if(c==16){uint8_t r=lens[idx-1];int rep=s.get(2)+3;while(rep--)lens[idx++]=r;}
                    else if(c==17){idx+=s.get(3)+3;}
                    else{idx+=s.get(7)+11;}
                }
                lit.n=hlit; for(int i=0;i<hlit;i++){lit.sym[i]=(uint16_t)i;lit.len[i]=lens[i];}
                dist.n=hdist; for(int i=0;i<hdist;i++){dist.sym[i]=(uint16_t)i;dist.len[i]=lens[hlit+i];}
            }
            while(s.ok){
                uint32_t sym=lit.decode(s);
                if(sym<256){s.put((uint8_t)sym);}
                else if(sym==256){break;}
                else{
                    int li=sym-257;
                    int len=lbase[li]+s.get(lext[li]);
                    uint32_t di=dist.decode(s);
                    int d=dbase[di]+s.get(dext[di]);
                    for(int i=0;i<len;i++) s.put(s.back(d));
                }
            }
        }
    }
    return s.ok?(int)s.dstPos:-1;
}
} // tinfl

// ── ZIP directory parser ──────────────────────────────────────────────────────
static vector<ZipEntry> zipList(const vector<uint8_t>& z) {
    vector<ZipEntry> entries;
    if (z.size() < 22) return entries;
    int eocd = -1;
    for (int i = (int)z.size()-22; i >= 0; i--)
        if (z[i]==0x50&&z[i+1]==0x4B&&z[i+2]==0x05&&z[i+3]==0x06){eocd=i;break;}
    if (eocd < 0) return entries;
    uint32_t cdOff = ru32(&z[eocd+16]);
    uint16_t cdNum = ru16(&z[eocd+8]);
    uint32_t pos = cdOff;
    for (int i = 0; i < cdNum && pos+46 <= z.size(); i++) {
        if (ru32(&z[pos]) != 0x02014B50) break;
        uint16_t method = ru16(&z[pos+10]);
        uint32_t csz  = ru32(&z[pos+20]);
        uint32_t usz  = ru32(&z[pos+24]);
        uint16_t flen = ru16(&z[pos+28]);
        uint16_t elen = ru16(&z[pos+30]);
        uint16_t clen = ru16(&z[pos+32]);
        uint32_t lhOff = ru32(&z[pos+42]);
        string name((char*)&z[pos+46], flen);
        uint32_t dataOff = lhOff + 30 + ru16(&z[lhOff+26]) + ru16(&z[lhOff+28]);
        entries.push_back({name, dataOff, csz, usz, method});
        pos += 46 + flen + elen + clen;
    }
    return entries;
}

static string zipExtract(const vector<uint8_t>& z, const ZipEntry& e) {
    if (e.method == 0) { 
        return string((char*)&z[e.offset], e.compSize);
    }
    if (e.method == 8) { 
        vector<uint8_t> out(e.uncompSize + 1);
        int r = tinfl::inflate(&z[e.offset], e.compSize, out.data(), out.size());
        if (r < 0) return {};
        return string((char*)out.data(), r);
    }
    return {};
}

// ── TINY XML PARSER ─────────────────────────────────────────────────────────
struct XmlNode {
    string tag;
    map<string,string> attr;
    vector<XmlNode> children;
    string text;
};

static string xmlAttrVal(const string& src, const string& name) {
    size_t p = 0;
    while ((p = src.find(name, p)) != string::npos) {
        size_t q = p + name.size();
        while (q < src.size() && src[q]==' ') q++;
        if (q < src.size() && src[q]=='=') {
            q++;
            while (q < src.size() && src[q]==' ') q++;
            if (q < src.size() && (src[q]=='"'||src[q]=='\'')) {
                char delim = src[q++];
                size_t end = src.find(delim, q);
                if (end != string::npos) return src.substr(q, end-q);
            }
        }
        p++;
    }
    return {};
}

static string xmlUnescape(const string& s) {
    string r; r.reserve(s.size());
    for (size_t i = 0; i < s.size(); ) {
        if (s[i]=='&') {
            size_t e = s.find(';', i);
            if (e != string::npos) {
                string ent = s.substr(i+1, e-i-1);
                if (ent=="amp")  {r+='&'; i=e+1; continue;}
                if (ent=="lt")   {r+='<'; i=e+1; continue;}
                if (ent=="gt")   {r+='>'; i=e+1; continue;}
                if (ent=="quot") {r+='"'; i=e+1; continue;}
                if (ent=="apos") {r+='\'';i=e+1; continue;}
                if (ent.size()>1&&ent[0]=='#') {
                    int code = (ent[1]=='x') ? stoi(ent.substr(2),nullptr,16) : stoi(ent.substr(1));
                    if (code<0x80){r+=(char)code;}
                    else if(code<0x800){r+=(char)(0xC0|(code>>6));r+=(char)(0x80|(code&0x3F));}
                    else{r+=(char)(0xE0|(code>>12));r+=(char)(0x80|((code>>6)&0x3F));r+=(char)(0x80|(code&0x3F));}
                    i=e+1; continue;
                }
            }
        }
        r+=s[i++];
    }
    return r;
}

enum TokType { TOK_OPEN, TOK_CLOSE, TOK_SELF, TOK_TEXT };
struct Token { TokType type; string raw; string tag; };

static vector<Token> xmlTokenize(const string& xml) {
    vector<Token> toks;
    size_t i = 0, n = xml.size();
    while (i < n) {
        if (xml[i] != '<') {
            size_t j = xml.find('<', i);
            string txt = xml.substr(i, (j==string::npos?n:j)-i);
            bool allws = true;
            for (char c : txt) if (!isspace((unsigned char)c)){allws=false;break;}
            if (!allws) toks.push_back({TOK_TEXT, txt, {}});
            i = (j==string::npos?n:j);
            continue;
        }
        if (i+1<n && xml[i+1]=='?') { size_t e=xml.find("?>",i+2); i=(e==string::npos?n:e+2); continue; }
        
        if (i+3<n && xml.substr(i,4)=="<!--") { size_t e=xml.find("-->",i+4); i=(e==string::npos?n:e+3); continue; }
        
        if (i+8<n && xml.substr(i,9)=="<![CDATA["){ size_t e=xml.find("]]>",i+9);
            if(e!=string::npos){toks.push_back({TOK_TEXT,xml.substr(i+9,e-i-9),{}});i=e+3;}else i=n; continue;}
        size_t j = xml.find('>', i);
        if (j == string::npos) break;
        string raw = xml.substr(i+1, j-i-1);
        i = j+1;
        if (!raw.empty() && raw[0]=='/') {
            string tag = raw.substr(1);
            auto col = tag.find(':'); if(col!=string::npos) tag=tag.substr(col+1);
            size_t sp=tag.find_first_of(" \t\r\n"); if(sp!=string::npos) tag=tag.substr(0,sp);
            toks.push_back({TOK_CLOSE,raw,tag});
        } else {
            bool self = (!raw.empty() && raw.back()=='/');
            if(self) raw.pop_back();
            size_t sp = raw.find_first_of(" \t\r\n:/");
            string tag = (sp==string::npos)?raw:raw.substr(0,sp);
            auto col = tag.find(':'); if(col!=string::npos) tag=tag.substr(col+1);
            toks.push_back({self?TOK_SELF:TOK_OPEN, raw, tag});
        }
    }
    return toks;
}

static XmlNode xmlBuild(const vector<Token>& toks, size_t& i) {
    XmlNode node;
    node.tag = toks[i].tag;
    const string& raw = toks[i].raw;
    size_t p = raw.find_first_of(" \t\r\n");
    while (p != string::npos && p < raw.size()) {
        while (p < raw.size() && isspace((unsigned char)raw[p])) p++;
        size_t eq = raw.find('=', p);
        if (eq == string::npos) break;
        string aname = raw.substr(p, eq-p);
        auto col = aname.find(':'); if(col!=string::npos) aname=aname.substr(col+1);
        while(!aname.empty()&&isspace((unsigned char)aname.back())) aname.pop_back();
        eq++;
        while (eq < raw.size() && isspace((unsigned char)raw[eq])) eq++;
        if (eq < raw.size() && (raw[eq]=='"'||raw[eq]=='\'')) {
            char d=raw[eq++];
            size_t e=raw.find(d,eq);
            if(e!=string::npos){node.attr[aname]=xmlUnescape(raw.substr(eq,e-eq));p=e+1;}
            else break;
        } else break;
    }
    if (toks[i].type == TOK_SELF) { i++; return node; }
    i++;
    while (i < toks.size()) {
        if (toks[i].type == TOK_CLOSE && toks[i].tag == node.tag) { i++; break; }
        if (toks[i].type == TOK_TEXT) { node.text += xmlUnescape(toks[i].raw); i++; }
        else { node.children.push_back(xmlBuild(toks, i)); }
    }
    return node;
}

static XmlNode parseXml(const string& xml) {
    auto toks = xmlTokenize(xml);
    if (toks.empty()) return {};
    size_t i = 0;
    while (i < toks.size() && toks[i].type != TOK_OPEN && toks[i].type != TOK_SELF) i++;
    if (i >= toks.size()) return {};
    return xmlBuild(toks, i);
}

static const XmlNode* findFirst(const XmlNode& n, const string& tag) {
    if (n.tag == tag) return &n;
    for (auto& c : n.children) { auto* r=findFirst(c,tag); if(r) return r; }
    return nullptr;
}
static vector<const XmlNode*> findAll(const XmlNode& n, const string& tag) {
    vector<const XmlNode*> r;
    if (n.tag == tag) r.push_back(&n);
    for (auto& c : n.children) { auto sub=findAll(c,tag); r.insert(r.end(),sub.begin(),sub.end()); }
    return r;
}
static string attr(const XmlNode& n, const string& a){ auto it=n.attr.find(a); return it!=n.attr.end()?it->second:""; }

// ─────────────────────────────────────────────────────────────────────────────
// EMU → CSS
// ─────────────────────────────────────────────────────────────────────────────
static double emuToPx(const string& s) {
    if (s.empty()) return 0;
    try { return stoll(s) * 96.0 / 914400.0; } catch(...){ return 0; }
}
static string px(double v){ return to_string((int)round(v))+"px"; }

static string color(const string& s){
    if(s.empty()) return "transparent";
    if(s.size()==6) return "#"+s;
    return "#"+s;
}

// ─────────────────────────────────────────────────────────────────────────────
// PPTX → HTML SLIDE RENDERER
// ─────────────────────────────────────────────────────────────────────────────
struct PptxFile {
    vector<uint8_t> raw;
    vector<ZipEntry> entries;
    string get(const string& path) {
        for (auto& e : entries)
            if (e.name == path || e.name == "/"+path)
                return zipExtract(raw, e);
        string lp = path; for(char&c:lp) c=tolower(c);
        for (auto& e : entries) {
            string le=e.name; for(char&c:le) c=tolower(c);
            if(le==lp||le=="/"+lp) return zipExtract(raw,e);
        }
        return {};
    }
    bool load(const wstring& path) {
        raw = readFile(path);
        if(raw.empty()) return false;
        entries = zipList(raw);
        return !entries.empty();
    }
};

struct ThemeColors {
    map<string,string> scheme; 
    bool loaded = false;
};

static string resolveFillColor(const XmlNode& fillContainer, const ThemeColors& tc) {
    auto* srgb = findFirst(fillContainer, "srgbClr");
    if (srgb) {
        string v = attr(*srgb,"val");
        if (!v.empty()) return "#"+v;
    }
    return {};
}

static string runStyle(const XmlNode& rPr, const ThemeColors& tc) {
    string css;
    string sz = attr(rPr,"sz");
    if(!sz.empty()) css += "font-size:"+to_string(stoi(sz)/2)+"pt;";
    string b = attr(rPr,"b");
    if(b=="1"||b=="true") css += "font-weight:bold;";
    string i = attr(rPr,"i");
    if(i=="1"||i=="true") css += "font-style:italic;";
    string u = attr(rPr,"u");
    if(!u.empty()&&u!="none") css += "text-decoration:underline;";
    auto* solidFill = findFirst(rPr,"solidFill");
    if(solidFill){
        string resolved = resolveFillColor(*solidFill, tc);
        if(!resolved.empty()) css += "color:"+resolved+";";
    }
    auto* latin = findFirst(rPr,"latin");
    if(latin){
        string typeface = attr(*latin,"typeface");
        if(!typeface.empty()&&typeface!="+mj-lt"&&typeface!="+mn-lt")
            css += "font-family:'"+typeface+"',sans-serif;";
    }
    return css;
}

static string renderShape(const XmlNode& sp, double slideW, double slideH) {
    auto* spPr = findFirst(sp,"spPr");
    if(!spPr) return {};
    auto* xfrm  = findFirst(*spPr,"xfrm");
    auto* prstGeom = findFirst(*spPr,"prstGeom");

    double x=0,y=0,w=0,h=0;
    if(xfrm){
        auto* off = findFirst(*xfrm,"off");
        auto* ext = findFirst(*xfrm,"ext");
        if(off){x=emuToPx(attr(*off,"x")); y=emuToPx(attr(*off,"y"));}
        if(ext){w=emuToPx(attr(*ext,"cx")); h=emuToPx(attr(*ext,"cy"));}
    }

    string shapeType = prstGeom ? attr(*prstGeom,"prst") : "rect";
    string bgCss;
    auto* solidFill = findFirst(*spPr,"solidFill");
    if(solidFill){
        auto* srgb = findFirst(*solidFill,"srgbClr");
        if(srgb) bgCss = "background:"+color(attr(*srgb,"val"))+";";
    }
    auto* noFill = findFirst(*spPr,"noFill");
    if(noFill) bgCss = "background:transparent;";

    string borderCss;
    auto* ln = findFirst(*spPr,"ln");
    if(ln){
        auto* lnNoFill = findFirst(*ln,"noFill");
        if(lnNoFill){ borderCss="border:none;"; }
        else {
            double lnW = 0;
            string lnWstr = attr(*ln,"w");
            if(!lnWstr.empty()) lnW = stoll(lnWstr)*96.0/914400.0;
            string lnColor = "transparent";
            auto* lnSolid = findFirst(*ln,"solidFill");
            if(lnSolid){
                auto* lsrgb = findFirst(*lnSolid,"srgbClr");
                if(lsrgb) lnColor = color(attr(*lsrgb,"val"));
            }
            borderCss = "border:"+to_string((int)max(1.0,lnW))+"px solid "+lnColor+";";
        }
    }

    auto* txBody = findFirst(sp,"txBody");
    string textHtml;
    ThemeColors tc; // Mock theme
    if(txBody){
        auto* bodyPr = findFirst(*txBody,"bodyPr");
        string anchor = bodyPr ? attr(*bodyPr,"anchor") : "ctr";
        string vAlign = (anchor=="t")?"flex-start":(anchor=="b")?"flex-end":"center";

        textHtml = "<div style=\"position:absolute;inset:0;display:flex;flex-direction:column;"
                   "justify-content:"+vAlign+";overflow:hidden;padding:6px 8px;box-sizing:border-box;\">";

        auto paras = findAll(*txBody,"p");
        for(auto* para : paras){
            auto* pPr = findFirst(*para,"pPr");
            string algn = pPr ? attr(*pPr,"algn") : "";
            string textAlign = (algn=="ctr")?"center":(algn=="r")?"right":(algn=="just")?"justify":"left";

            textHtml += "<p style=\"margin:0;padding:0;text-align:"+textAlign+";line-height:1.2;\">";
            auto runs = findAll(*para,"r");
            for(auto* run : runs){
                auto* rPr = findFirst(*run,"rPr");
                string style = rPr ? runStyle(*rPr, tc) : "";
                auto* t = findFirst(*run,"t");
                string txt = t ? t->text : "";
                string esc;
                for(char c:txt){
                    if(c=='<') esc+="&lt;";
                    else if(c=='>') esc+="&gt;";
                    else if(c=='&') esc+="&amp;";
                    else esc+=c;
                }
                textHtml += "<span contenteditable=\"true\" style=\""+style+" outline: none;\">"+esc+"</span>";
            }
            textHtml += "</p>";
        }
        textHtml += "</div>";
    }

    string rotCss;
    if(xfrm){
        string rot = attr(*xfrm,"rot");
        if(!rot.empty()){
            double deg = stoll(rot)/60000.0;
            rotCss = "transform:rotate("+to_string(deg)+"deg);transform-origin:center center;";
        }
    }

    string div = "<div style=\"position:absolute;"
        "left:"+px(x)+";top:"+px(y)+";width:"+px(w)+";height:"+px(h)+";"
        +bgCss+borderCss+rotCss
        +"overflow:hidden;"
        +"box-sizing:border-box;\">"
        + textHtml +
        "</div>\n";
    return div;
}

static string renderSlideHtml(const string& slideXml, int slideNum, int totalSlides, double slideWpx, double slideHpx) {
    XmlNode slideDoc = parseXml(slideXml);
    string bgColor = "#FFFFFF";

    string shapesHtml;
    auto shapes = findAll(slideDoc,"sp");
    for(auto* sp : shapes)
        shapesHtml += renderShape(*sp, slideWpx, slideHpx);

    string w = to_string((int)slideWpx), h = to_string((int)slideHpx);

    return R"(<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
*{margin:0;padding:0;box-sizing:border-box}
html,body{width:100%;height:100%;background:#1a1a2e;display:flex;align-items:center;justify-content:center;font-family:Calibri,Segoe UI,Arial,sans-serif;}
#slide{position:relative;width:)" + w + R"(px;height:)" + h + R"(px;background:)" + bgColor + R"(;overflow:hidden;box-shadow:0 8px 40px rgba(0,0,0,.5);}
p{min-height:1em}
[contenteditable="true"]:hover { background: rgba(255,255,255,0.1); cursor: text; }
</style></head><body>
<div id="slide">
)" + shapesHtml + R"(
</div>
<script>
window.chrome && chrome.webview && chrome.webview.postMessage(JSON.stringify({type:'ready',slide:)" 
+ to_string(slideNum) + R"(,total:)" + to_string(totalSlides) + R"(}));

// Listen for Ctrl+S to save edits
document.addEventListener('keydown', function(e) {
    if ((e.ctrlKey || e.metaKey) && e.key === 's') {
        e.preventDefault();
        window.chrome && chrome.webview && chrome.webview.postMessage(JSON.stringify({type:'save'}));
        
        let slide = document.getElementById('slide');
        slide.style.opacity = '0.7';
        setTimeout(() => slide.style.opacity = '1', 200);
    }
});
</script>
</body></html>)";
}

// ─────────────────────────────────────────────────────────────────────────────
// APPLICATION STATE
// ─────────────────────────────────────────────────────────────────────────────
static HWND g_hwnd = nullptr;
static ComPtr<ICoreWebView2Controller> g_controller;
static ComPtr<ICoreWebView2> g_webview;
static PptxFile g_pptx;
static int g_currentSlide = 0;
static int g_totalSlides  = 0;
static bool g_loaded = false;

static double g_slideWpx = 960.0;
static double g_slideHpx = 540.0;

static void showSlide(int idx) {
    if(!g_loaded || !g_webview) return;
    if(idx<0) idx=0;
    if(idx>=g_totalSlides) idx=g_totalSlides-1;
    g_currentSlide = idx;

    string slideXml  = g_pptx.get("ppt/slides/slide"+to_string(idx+1)+".xml");
    string html = renderSlideHtml(slideXml, idx+1, g_totalSlides, g_slideWpx, g_slideHpx);

    g_webview->NavigateToString(wstring(html.begin(),html.end()).c_str());
    wstring title = L"PPTX Viewer  —  Slide " + to_wstring(idx+1) + L" / " + to_wstring(g_totalSlides);
    SetWindowTextW(g_hwnd, title.c_str());
}

static void openFile(wstring path = L"") {
    if(path.empty()){
        OPENFILENAMEW ofn = {};
        wchar_t buf[MAX_PATH] = {};
        ofn.lStructSize = sizeof(ofn);
        ofn.hwndOwner   = g_hwnd;
        ofn.lpstrFilter = L"PowerPoint Files\0*.pptx;*.ppt\0All Files\0*.*\0";
        ofn.lpstrFile   = buf;
        ofn.nMaxFile    = MAX_PATH;
        ofn.Flags       = OFN_FILEMUSTEXIST;
        if(!GetOpenFileNameW(&ofn)) return;
        path = buf;
    }
    if(!g_pptx.load(path)){
        MessageBoxW(g_hwnd, L"Failed to open the PPTX file.", L"Error", MB_ICONERROR);
        return;
    }
    g_totalSlides = 0;
    for(auto& e : g_pptx.entries){
        if(e.name.find("ppt/slides/slide")!=string::npos && e.name.find(".xml")!=string::npos && e.name.find(".rels")==string::npos)
            g_totalSlides++;
    }
    if(g_totalSlides==0) return;
    g_loaded=true;
    showSlide(0);
}

static void resizeWebView() {
    if(!g_controller) return;
    RECT rc; GetClientRect(g_hwnd,&rc);
    g_controller->put_Bounds({rc.left,rc.top+40,rc.right,rc.bottom});
}

static void initWebView() {
    wchar_t tempPath[MAX_PATH];
    GetTempPathW(MAX_PATH,tempPath);

    CreateCoreWebView2EnvironmentWithOptions(nullptr,tempPath,nullptr,
        Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>(
            [](HRESULT hr, ICoreWebView2Environment* env) -> HRESULT {
                if(FAILED(hr)||!env) return hr;
                env->CreateCoreWebView2Controller(g_hwnd,
                    Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
                        [](HRESULT hr, ICoreWebView2Controller* ctrl) -> HRESULT {
                            if(FAILED(hr)||!ctrl) return hr;
                            g_controller = ctrl;
                            ctrl->get_CoreWebView2(&g_webview);

                            ComPtr<ICoreWebView2Settings> settings;
                            g_webview->get_Settings(&settings);
                            if(settings) settings->put_IsScriptEnabled(TRUE);

                            g_webview->add_WebMessageReceived(
                                Callback<ICoreWebView2WebMessageReceivedEventHandler>(
                                    [](ICoreWebView2* wv, ICoreWebView2WebMessageReceivedEventArgs* args) -> HRESULT {
                                        LPWSTR msg; args->TryGetWebMessageAsString(&msg);
                                        if(msg) {
                                            wstring wmsg(msg);
                                            if(wmsg.find(L"\"type\":\"save\"") != wstring::npos) {
                                                MessageBoxW(g_hwnd, L"Save signal received! (XML Patching logic needs to be added here)", L"Save Command", MB_OK | MB_ICONINFORMATION);
                                            }
                                            CoTaskMemFree(msg);
                                        }
                                        return S_OK;
                                    }).Get(), nullptr);

                            resizeWebView();
                            wstring welcome = LR"(<!DOCTYPE html><html><body style="background:#1a1a2e;color:#fff;display:flex;align-items:center;justify-content:center;height:100vh;font-family:sans-serif;"><h1>PPTX Viewer</h1></body></html>)";
                            g_webview->NavigateToString(welcome.c_str());
                            return S_OK;
                        }).Get());
                return S_OK;
            }).Get());
}

static void paintToolbar(HWND hwnd) {
    PAINTSTRUCT ps; HDC hdc=BeginPaint(hwnd,&ps);
    RECT rc; GetClientRect(hwnd,&rc); rc.bottom=40;
    HBRUSH bg=CreateSolidBrush(RGB(30,30,50)); FillRect(hdc,&rc,bg); DeleteObject(bg);
    SetBkMode(hdc,TRANSPARENT); SetTextColor(hdc,RGB(220,220,240));
    RECT br={8,6,120,34};
    DrawTextW(hdc,L"Open",-1,&br,DT_CENTER|DT_VCENTER|DT_SINGLELINE);
    EndPaint(hwnd,&ps);
}

LRESULT CALLBACK WndProc(HWND hwnd,UINT msg,WPARAM wp,LPARAM lp){
    switch(msg){
    case WM_PAINT: paintToolbar(hwnd); return 0;
    case WM_SIZE:  resizeWebView(); return 0;
    case WM_LBUTTONDOWN:
        if(GET_X_LPARAM(lp)<120 && GET_Y_LPARAM(lp)<40) openFile();
        return 0;
    case WM_KEYDOWN:
        if(wp==VK_RIGHT) showSlide(g_currentSlide+1);
        if(wp==VK_LEFT) showSlide(g_currentSlide-1);
        return 0;
    case WM_DESTROY: PostQuitMessage(0); return 0;
    }
    return DefWindowProcW(hwnd,msg,wp,lp);
}

int WINAPI wWinMain(HINSTANCE hInst,HINSTANCE,LPWSTR,int nCmdShow){
    CoInitializeEx(nullptr,COINIT_APARTMENTTHREADED);
    WNDCLASSEXW wc={sizeof(wc)};
    wc.lpfnWndProc   = WndProc;
    wc.hInstance     = hInst;
    wc.hCursor       = LoadCursor(nullptr,IDC_ARROW);
    wc.lpszClassName = L"PptxViewerWnd";
    RegisterClassExW(&wc);

    g_hwnd = CreateWindowExW(0,L"PptxViewerWnd",L"PPTX Viewer",WS_OVERLAPPEDWINDOW,CW_USEDEFAULT,CW_USEDEFAULT,1280,760,nullptr,nullptr,hInst,nullptr);
    ShowWindow(g_hwnd,nCmdShow);
    initWebView();

    MSG msg={};
    while(GetMessageW(&msg,nullptr,0,0)){ TranslateMessage(&msg); DispatchMessageW(&msg); }
    CoUninitialize();
    return (int)msg.wParam;
}
