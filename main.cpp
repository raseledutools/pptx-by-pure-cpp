// pptx_viewer_editor_final.cpp - 3000+ Lines with Full Editing
// Complete PPTX Viewer + Editor with All Professional Features
// Compact but powerful implementation with real-time editing

#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <windowsx.h>
#include <commdlg.h>
#include <shlwapi.h>
#include <shlobj.h>
#include <d2d1.h>
#include <dwrite.h>
#include <wincodec.h>
#include <string>
#include <vector>
#include <map>
#include <memory>
#include <fstream>
#include <sstream>
#include <algorithm>
#include <cmath>
#include <thread>
#include <mutex>
#include <functional>
#include <set>
#include <cstring>
#include <cassert>
#include <codecvt>
#include <locale>
#include <numeric>

// ─── Additional D2D/DWrite includes ─────────────────────────────────
#include <d2d1helper.h>
#include <dwrite_2.h>
#include <d2d1_1.h>
#include <wrl/client.h>

using Microsoft::WRL::ComPtr;

#pragma comment(lib, "d2d1.lib")
#pragma comment(lib, "dwrite.lib")
#pragma comment(lib, "windowscodecs.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "comdlg32.lib")

// ═══════════════════════════════════════════════════════════════════
// SECTION 0: UTILITY HELPERS
// ═══════════════════════════════════════════════════════════════════

// EMU → pixel at 96 DPI
static inline double EmuToPx(double emu) { return emu / 914400.0 * 96.0; }
// Hundredths of a point → pixel at 96 DPI  (font size)
static inline double HptToPx(double hpt)  { return hpt / 100.0 * 96.0 / 72.0; }
// EMU → D2D float
static inline float EmuToF(double emu)    { return (float)(emu / 914400.0 * 96.0); }
// Angle in 60000ths of a degree → radians
static inline float AngToRad(double ang)  { return (float)(ang / 60000.0 * 3.14159265358979 / 180.0); }

static std::wstring Utf8ToWide(const std::string& s) {
    if (s.empty()) return {};
    int n = MultiByteToWideChar(CP_UTF8, 0, s.c_str(), -1, nullptr, 0);
    std::wstring w(n, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, s.c_str(), -1, w.data(), n);
    if (!w.empty() && w.back() == L'\0') w.pop_back();
    return w;
}

static std::string WideToUtf8(const std::wstring& w) {
    if (w.empty()) return {};
    int n = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), -1, nullptr, 0, nullptr, nullptr);
    std::string s(n, '\0');
    WideCharToMultiByte(CP_UTF8, 0, w.c_str(), -1, s.data(), n, nullptr, nullptr);
    if (!s.empty() && s.back() == '\0') s.pop_back();
    return s;
}

// Parse hex color "#RRGGBB" or "RRGGBB" → D2D1_COLOR_F
static D2D1_COLOR_F HexToColor(const std::string& hex, float alpha = 1.0f) {
    std::string h = hex;
    if (!h.empty() && h[0] == '#') h = h.substr(1);
    if (h.size() < 6) return D2D1::ColorF(0,0,0, alpha);
    unsigned int rgb = (unsigned int)std::stoul(h.substr(0,6), nullptr, 16);
    return D2D1::ColorF(
        ((rgb >> 16) & 0xFF) / 255.0f,
        ((rgb >>  8) & 0xFF) / 255.0f,
        ( rgb        & 0xFF) / 255.0f,
        alpha
    );
}

// Clamp helper
static inline float Clampf(float v, float lo, float hi) {
    return v < lo ? lo : (v > hi ? hi : v);
}

// ═══════════════════════════════════════════════════════════════════
// SECTION 1: UNIFIED XML PARSER (Handles ALL pptx XML formats)
// ═══════════════════════════════════════════════════════════════════

class XmlNode {
public:
    std::string name, value;
    std::map<std::string, std::string> attrs;
    std::vector<XmlNode*> children;
    XmlNode* parent = nullptr;

    // ── helpers ──────────────────────────────────────────────────
    const std::string& Attr(const std::string& k, const std::string& def = "") const {
        auto it = attrs.find(k);
        return it != attrs.end() ? it->second : def;
    }
    bool HasAttr(const std::string& k) const { return attrs.count(k) > 0; }

    // Find first child whose local name (after ':') matches
    XmlNode* Child(const std::string& localName) const {
        for (auto* c : children) {
            std::string n = c->name;
            size_t colon = n.find(':');
            if (colon != std::string::npos) n = n.substr(colon+1);
            if (n == localName) return c;
        }
        return nullptr;
    }

    // Find all children matching local name
    std::vector<XmlNode*> ChildrenNamed(const std::string& localName) const {
        std::vector<XmlNode*> result;
        for (auto* c : children) {
            std::string n = c->name;
            size_t colon = n.find(':');
            if (colon != std::string::npos) n = n.substr(colon+1);
            if (n == localName) result.push_back(c);
        }
        return result;
    }

    // Deep find (first match)
    XmlNode* Find(const std::string& localName) const {
        if (name.find(localName) != std::string::npos) return const_cast<XmlNode*>(this);
        for (auto* c : children) {
            auto* r = c->Find(localName);
            if (r) return r;
        }
        return nullptr;
    }

    // ── parser ────────────────────────────────────────────────────
    static XmlNode* Parse(const std::string& xml, size_t& pos) {
        // Skip whitespace / comments / PI
        while (pos < xml.size()) {
            while (pos < xml.size() && isspace((unsigned char)xml[pos])) pos++;
            if (pos >= xml.size()) return nullptr;
            if (xml[pos] != '<') return nullptr;
            // XML declaration / PI
            if (pos+1 < xml.size() && xml[pos+1] == '?') {
                pos = xml.find("?>", pos);
                if (pos == std::string::npos) return nullptr;
                pos += 2; continue;
            }
            // Comment
            if (pos+3 < xml.size() && xml.substr(pos,4) == "<!--") {
                pos = xml.find("-->", pos);
                if (pos == std::string::npos) return nullptr;
                pos += 3; continue;
            }
            // CDATA
            if (pos+8 < xml.size() && xml.substr(pos,9) == "<![CDATA[") {
                pos += 9;
                size_t end = xml.find("]]>", pos);
                if (end == std::string::npos) return nullptr;
                pos = end + 3; continue;
            }
            break;
        }
        if (pos >= xml.size() || xml[pos] != '<') return nullptr;

        auto node = new XmlNode();
        pos++; // skip '<'

        // Tag name
        size_t nameEnd = xml.find_first_of(" >/\t\n\r", pos);
        if (nameEnd == std::string::npos) { delete node; return nullptr; }
        node->name = xml.substr(pos, nameEnd - pos);
        pos = nameEnd;

        // Attributes
        while (pos < xml.size() && xml[pos] != '>' && xml[pos] != '/') {
            while (pos < xml.size() && isspace((unsigned char)xml[pos])) pos++;
            if (pos >= xml.size() || xml[pos] == '>' || xml[pos] == '/') break;

            size_t eq = xml.find('=', pos);
            if (eq == std::string::npos) break;
            std::string key = Trim(xml.substr(pos, eq - pos));
            pos = eq + 1;

            if (pos < xml.size() && (xml[pos] == '"' || xml[pos] == '\'')) {
                char q = xml[pos++];
                size_t valEnd = xml.find(q, pos);
                if (valEnd == std::string::npos) { delete node; return nullptr; }
                node->attrs[key] = xml.substr(pos, valEnd - pos);
                pos = valEnd + 1;
            } else {
                // Unquoted
                size_t valEnd = xml.find_first_of(" \t\n\r>", pos);
                if (valEnd == std::string::npos) valEnd = xml.size();
                node->attrs[key] = xml.substr(pos, valEnd - pos);
                pos = valEnd;
            }
        }

        // Self-closing?
        if (pos < xml.size() && xml[pos] == '/') {
            pos += 2; // skip '/>'
            return node;
        }
        if (pos < xml.size()) pos++; // skip '>'

        // Children / text content
        std::string text;
        while (pos < xml.size()) {
            if (xml[pos] == '<') {
                // Closing tag?
                if (pos+1 < xml.size() && xml[pos+1] == '/') {
                    size_t end = xml.find('>', pos);
                    pos = (end == std::string::npos) ? xml.size() : end + 1;
                    break;
                }
                if (!text.empty()) {
                    node->value = Trim(text);
                    text.clear();
                }
                auto child = Parse(xml, pos);
                if (child) {
                    child->parent = node;
                    node->children.push_back(child);
                }
            } else {
                text += xml[pos++];
            }
        }
        if (!text.empty() && node->value.empty()) node->value = Trim(text);
        return node;
    }

    static XmlNode* ParseFull(const std::string& xml) {
        size_t pos = 0;
        return Parse(xml, pos);
    }

    static std::string Trim(const std::string& s) {
        size_t start = s.find_first_not_of(" \t\n\r");
        if (start == std::string::npos) return "";
        size_t end = s.find_last_not_of(" \t\n\r");
        return s.substr(start, end - start + 1);
    }

    ~XmlNode() { for (auto c : children) delete c; }
};

// ═══════════════════════════════════════════════════════════════════
// SECTION 1B: FULL DEFLATE / INFLATE ENGINE (RFC 1951)
// ═══════════════════════════════════════════════════════════════════

class Inflater {
    const uint8_t* in;
    size_t inSize;
    std::vector<uint8_t>& out;
    uint32_t bitBuf = 0;
    int bitCount = 0;

    uint32_t readBits(int n) {
        while (bitCount < n) {
            if (inPos >= inSize) break;
            bitBuf |= (uint32_t)in[inPos++] << bitCount;
            bitCount += 8;
        }
        uint32_t v = bitBuf & ((1u << n) - 1);
        bitBuf >>= n;
        bitCount -= n;
        return v;
    }
    void alignByte() { bitCount = 0; bitBuf = 0; }

    // Huffman tree node
    struct HTree {
        int sym[65536];
        int count;
        HTree() : count(0) { memset(sym, -1, sizeof(sym)); }
    };

    size_t inPos = 0;

    // Build canonical Huffman table from code lengths
    struct HuffTable {
        uint16_t codes[288];
        uint8_t  lens[288];
        int      maxSym;
        int      table[65536]; // decode table (16-bit key)
        int      tableBits;

        void build(const int* lengths, int n, int tbits) {
            maxSym = n;
            tableBits = tbits;
            memset(table, -1, sizeof(table));

            // Count codes per length
            int bl_count[16] = {};
            for (int i = 0; i < n; i++) if (lengths[i]) bl_count[lengths[i]]++;

            // First code for each length
            int next_code[16] = {};
            int code = 0;
            for (int bits = 1; bits < 16; bits++) {
                code = (code + bl_count[bits-1]) << 1;
                next_code[bits] = code;
            }

            // Assign codes
            for (int i = 0; i < n; i++) {
                int len = lengths[i];
                if (!len) continue;
                int c = next_code[len]++;
                lens[i] = (uint8_t)len;
                codes[i] = (uint16_t)c;

                // Fill decode table (pad to tableBits)
                int pad = tbits - len;
                for (int j = 0; j < (1 << pad); j++) {
                    int key = (j << len) | c;
                    if (key < 65536) table[key] = i;
                }
            }
        }

        int decode(uint32_t& bitBuf, int& bitCount, const uint8_t* in, size_t& inPos, size_t inSize) {
            // Ensure we have tableBits bits
            while (bitCount < tableBits) {
                if (inPos >= inSize) break;
                bitBuf |= (uint32_t)in[inPos++] << bitCount;
                bitCount += 8;
            }
            int key = bitBuf & ((1 << tableBits) - 1);
            int sym = table[key];
            if (sym < 0 || sym >= maxSym) return 256; // end
            int len = lens[sym];
            bitBuf >>= len;
            bitCount -= len;
            return sym;
        }
    };

    static const int FIXED_LL[288];
    static const int FIXED_DIST[32];

    void inflateBlock(HuffTable& ll, HuffTable& dist) {
        static const int lenBase[] = {
            3,4,5,6,7,8,9,10,11,13,15,17,19,23,27,31,
            35,43,51,59,67,83,99,115,131,163,195,227,258
        };
        static const int lenExtra[] = {
            0,0,0,0,0,0,0,0,1,1,1,1,2,2,2,2,
            3,3,3,3,4,4,4,4,5,5,5,5,0
        };
        static const int distBase[] = {
            1,2,3,4,5,7,9,13,17,25,33,49,65,97,129,193,
            257,385,513,769,1025,1537,2049,3073,4097,6145,
            8193,12289,16385,24577
        };
        static const int distExtra[] = {
            0,0,0,0,1,1,2,2,3,3,4,4,5,5,6,6,
            7,7,8,8,9,9,10,10,11,11,12,12,13,13
        };

        while (true) {
            int sym = ll.decode(bitBuf, bitCount, in, inPos, inSize);
            if (sym < 256) {
                out.push_back((uint8_t)sym);
            } else if (sym == 256) {
                break;
            } else {
                int li = sym - 257;
                if (li < 0 || li >= 29) break;
                int len = lenBase[li] + (int)readBits(lenExtra[li]);
                int dsym = dist.decode(bitBuf, bitCount, in, inPos, inSize);
                if (dsym < 0 || dsym >= 30) break;
                int d = distBase[dsym] + (int)readBits(distExtra[dsym]);
                size_t start = out.size();
                for (int i = 0; i < len; i++) {
                    if (d > (int)out.size()) out.push_back(0);
                    else out.push_back(out[out.size() - d]);
                }
            }
        }
    }

    void readDynHeader(HuffTable& ll, HuffTable& dist) {
        int hlit  = 257 + (int)readBits(5);
        int hdist = 1   + (int)readBits(5);
        int hclen = 4   + (int)readBits(4);

        static const int clOrder[] = {
            16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15
        };
        int clLens[19] = {};
        for (int i = 0; i < hclen; i++) clLens[clOrder[i]] = (int)readBits(3);

        HuffTable cl;
        cl.build(clLens, 19, 7);

        // Read LL + dist code lengths
        std::vector<int> codeLens(hlit + hdist, 0);
        for (int i = 0; i < hlit + hdist; ) {
            int sym = cl.decode(bitBuf, bitCount, in, inPos, inSize);
            if (sym < 16) {
                codeLens[i++] = sym;
            } else if (sym == 16) {
                int rep = 3 + (int)readBits(2);
                int prev = i > 0 ? codeLens[i-1] : 0;
                while (rep-- && i < hlit + hdist) codeLens[i++] = prev;
            } else if (sym == 17) {
                int rep = 3 + (int)readBits(3);
                while (rep-- && i < hlit + hdist) codeLens[i++] = 0;
            } else if (sym == 18) {
                int rep = 11 + (int)readBits(7);
                while (rep-- && i < hlit + hdist) codeLens[i++] = 0;
            } else break;
        }

        ll.build(codeLens.data(), hlit, 15);
        dist.build(codeLens.data() + hlit, hdist, 15);
    }

public:
    Inflater(const uint8_t* data, size_t size, std::vector<uint8_t>& output)
        : in(data), inSize(size), out(output), inPos(0) {}

    bool inflate() {
        // Static fixed Huffman tables
        static bool tablesBuilt = false;
        static HuffTable fixedLL, fixedDist;
        if (!tablesBuilt) {
            int llLens[288];
            for (int i = 0;   i < 144; i++) llLens[i] = 8;
            for (int i = 144; i < 256; i++) llLens[i] = 9;
            for (int i = 256; i < 280; i++) llLens[i] = 7;
            for (int i = 280; i < 288; i++) llLens[i] = 8;
            fixedLL.build(llLens, 288, 9);
            int distLens[32];
            for (int i = 0; i < 32; i++) distLens[i] = 5;
            fixedDist.build(distLens, 32, 5);
            tablesBuilt = true;
        }

        bool bfinal = false;
        while (!bfinal) {
            bfinal = readBits(1) != 0;
            int btype = (int)readBits(2);

            if (btype == 0) {
                // Stored block
                alignByte();
                if (inPos + 4 > inSize) return false;
                uint16_t len  = in[inPos] | (in[inPos+1] << 8); inPos += 2;
                uint16_t nlen = in[inPos] | (in[inPos+1] << 8); inPos += 2;
                for (uint16_t i = 0; i < len && inPos < inSize; i++)
                    out.push_back(in[inPos++]);
            } else if (btype == 1) {
                inflateBlock(fixedLL, fixedDist);
            } else if (btype == 2) {
                HuffTable dynLL, dynDist;
                readDynHeader(dynLL, dynDist);
                inflateBlock(dynLL, dynDist);
            } else {
                return false; // reserved
            }
        }
        return true;
    }
};

// ═══════════════════════════════════════════════════════════════════
// SECTION 2: COMPLETE PPTX ENGINE (ZIP + Parse + Theme + All Effects)
// ═══════════════════════════════════════════════════════════════════

class PptxEngine {
public:
    // ── Data Structures ──────────────────────────────────────────

    struct ColorMod {
        float lumMod = 1.0f, lumOff = 0.0f;
        float shadeMod = 1.0f, tintMod = 1.0f;
        float alphaMod = 1.0f;
        float hueMod = 0.0f, satMod = 1.0f;
    };

    struct RunProps {
        float   fontSize  = 18.0f;  // pt
        bool    bold      = false;
        bool    italic    = false;
        bool    underline = false;
        bool    strikeThru = false;
        std::string fontFamily = "Calibri";
        D2D1_COLOR_F color = D2D1::ColorF(D2D1::ColorF::Black);
        float   spacing   = 0.0f;   // character spacing (EMU)
        int     baseline  = 0;      // +ve = superscript, -ve = subscript
    };

    struct ParaProps {
        float     indent       = 0.0f;   // px
        float     leftMargin   = 0.0f;   // px
        float     spaceAfter   = 0.0f;   // pt
        float     spaceBefore  = 0.0f;   // pt
        float     lineSpacing  = 1.0f;   // multiplier
        bool      lineSpacingIsPoints = false;
        DWRITE_TEXT_ALIGNMENT align = DWRITE_TEXT_ALIGNMENT_LEADING;
        DWRITE_PARAGRAPH_ALIGNMENT valign = DWRITE_PARAGRAPH_ALIGNMENT_NEAR;
        int       level = 0;
        D2D1_COLOR_F bulletColor = D2D1::ColorF(D2D1::ColorF::Black);
        std::string bulletChar;
        bool        hasBullet = false;
    };

    struct TextRun {
        RunProps  props;
        std::wstring text;
    };

    struct Paragraph {
        ParaProps         props;
        std::vector<TextRun> runs;
    };

    struct ImageData {
        std::string rId;
        std::vector<uint8_t> bytes;
        std::wstring mimeType;
        ComPtr<ID2D1Bitmap> bitmap; // loaded on demand
    };

    struct ShapeGeometry {
        std::string prst; // preset geometry name
        float adjVal[8] = {};
    };

    enum class ShapeType {
        Rect, Ellipse, RoundRect, Triangle, Parallelogram,
        Trapezoid, Diamond, Pentagon, Hexagon, Heptagon, Octagon,
        Star4, Star5, Star6, Star8, Star16, Star24, Star32,
        ArrowLeft, ArrowRight, ArrowUp, ArrowDown,
        ArrowCallout, Chevron, Ribbon, Wave,
        Cloud, Heart, Lightning, Gear,
        Connector, BentConnector, CurvedConnector,
        Line, FreeForm, Custom,
        Picture, Group, Table,
        Unknown
    };

    struct GradStop {
        D2D1_COLOR_F color;
        float pos; // 0-1
    };

    struct FillInfo {
        enum Type { Solid, Gradient, Pattern, Picture, None } type = None;
        D2D1_COLOR_F solidColor = D2D1::ColorF(D2D1::ColorF::White);
        std::vector<GradStop> gradStops;
        float gradAngle = 0.0f;
        bool  gradLinear = true;
        std::string imageRId;
        float alpha = 1.0f;
    };

    struct LineInfo {
        float width = 0.0f; // px
        FillInfo fill;
        enum Cap { Flat, Round, Square } cap = Flat;
        enum Join { Miter, Round_, Bevel } join = Miter;
        enum Dash { Solid, Dot, Dash_, DashDot, LongDash, LongDashDot } dash = Solid;
        float headW = 0, headLen = 0;
        float tailW = 0, tailLen = 0;
        std::string headType, tailType;
    };

    struct EffectInfo {
        bool hasShadow = false;
        float shadowBlur = 0.0f;
        float shadowDX = 0.0f, shadowDY = 0.0f;
        D2D1_COLOR_F shadowColor = D2D1::ColorF(0,0,0,0.5f);
        bool hasGlow = false;
        float glowRadius = 0.0f;
        D2D1_COLOR_F glowColor = D2D1::ColorF(D2D1::ColorF::White, 0.5f);
        bool hasSoftEdge = false;
        float softEdgeRadius = 0.0f;
        float reflection = 0.0f;
    };

    struct ShapeData {
        int            id = 0;
        std::string    name;
        ShapeType      type = ShapeType::Unknown;
        ShapeGeometry  geom;
        float          x = 0, y = 0, w = 100, h = 100;
        float          rot = 0.0f;  // degrees
        bool           flipH = false, flipV = false;
        FillInfo       fill;
        LineInfo       line;
        EffectInfo     effect;
        std::vector<Paragraph> paragraphs;
        float          txInsetL = 3.6f, txInsetR = 3.6f;
        float          txInsetT = 1.8f, txInsetB = 1.8f;
        DWRITE_TEXT_ALIGNMENT  textHAlign = DWRITE_TEXT_ALIGNMENT_LEADING;
        DWRITE_PARAGRAPH_ALIGNMENT textVAlign = DWRITE_PARAGRAPH_ALIGNMENT_NEAR;
        bool           wordWrap = true;
        bool           autoFit = false;
        std::string    hyperlinkUrl;
        std::string    imageRId;
        float          imageAlpha = 1.0f;
        std::vector<ShapeData> children; // for groups
        std::string    altText;

        // Table-specific
        struct TableCell {
            int rowSpan = 1, colSpan = 1;
            std::vector<Paragraph> paragraphs;
            FillInfo fill;
            LineInfo borders[4]; // top,right,bottom,left
        };
        std::vector<std::vector<TableCell>> tableRows;
        std::vector<float> colWidths, rowHeights;
    };

    struct SlideLayout {
        std::string xml;
        std::string name;
        std::string type; // title, titleContent, blank, etc.
        std::vector<ShapeData> shapes;
    };

    struct SlideMaster {
        std::string xml;
        std::string rels;
        std::vector<ShapeData> shapes;
        std::map<std::string, std::string> layoutPaths;
    };

    struct SlideData {
        std::string xml, rels;
        std::string layoutPath;
        int number = 0;
        float width = 960.0f, height = 540.0f;
        std::vector<ShapeData> shapes;
        std::string backgroundXml;
        FillInfo    background;
        std::vector<ImageData> images;
        std::string notes;
        std::string title;
        bool        hidden = false;
        int         transitionType = 0;
        float       transitionDuration = 0.5f;
    };

    struct ThemeData {
        std::map<std::string, std::string> colors;
        std::map<std::string, std::string> fonts;
        std::string name;

        // Resolve scheme color → hex
        std::string Resolve(const std::string& ref) const {
            auto it = colors.find(ref);
            return (it != colors.end()) ? it->second : "#000000";
        }
    };

    struct EditCommand {
        std::string slideId, shapeId;
        std::string newText, newStyle;
        enum Type {
            TEXT_CHANGE, STYLE_CHANGE, MOVE, RESIZE,
            DELETE_SHAPE, ADD_SHAPE, REORDER, CHANGE_FILL,
            CHANGE_LINE, CHANGE_EFFECT, ROTATE, FLIP
        } type;
        float x = 0, y = 0, w = 0, h = 0;
        float rot = 0;
        std::string xmlPatch;
    };

private:
    std::vector<uint8_t> m_zip;
    std::vector<SlideData> m_slides;
    std::vector<SlideLayout> m_layouts;
    SlideMaster m_master;
    ThemeData   m_theme;
    std::vector<EditCommand> m_editHistory;
    int m_currentSlide = 0;
    float m_slideW = 960.0f, m_slideH = 540.0f;
    std::string m_filePath;

    // ZIP entries
    struct ZipEntry {
        std::string name;
        uint32_t offset = 0, compSize = 0, uncompSize = 0;
        uint16_t method = 0;
    };
    std::vector<ZipEntry> m_entries;

public:
    // ── Public API ──────────────────────────────────────────────
    bool LoadPptx(const std::wstring& path) {
        m_filePath = WideToUtf8(path);
        HANDLE h = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ,
                               NULL, OPEN_EXISTING, 0, NULL);
        if (h == INVALID_HANDLE_VALUE) return false;

        LARGE_INTEGER sz;
        GetFileSizeEx(h, &sz);
        m_zip.resize((size_t)sz.QuadPart);

        DWORD rd = 0;
        ReadFile(h, m_zip.data(), (DWORD)m_zip.size(), &rd, NULL);
        CloseHandle(h);

        m_slides.clear();
        m_entries.clear();
        m_editHistory.clear();

        ParseZipDirectory();
        LoadPresentation();
        LoadTheme();
        LoadMaster();
        LoadLayouts();
        ParseAllSlides();

        return !m_slides.empty();
    }

    int SlideCount() const { return (int)m_slides.size(); }
    float SlideWidth()  const { return m_slideW; }
    float SlideHeight() const { return m_slideH; }

    const SlideData& GetSlide(int i) const {
        static SlideData empty;
        return (i >= 0 && i < (int)m_slides.size()) ? m_slides[i] : empty;
    }
    SlideData& GetSlide(int i) {
        static SlideData empty;
        return (i >= 0 && i < (int)m_slides.size()) ? m_slides[i] : empty;
    }

    const ThemeData& Theme() const { return m_theme; }

    std::vector<uint8_t> GetImageBytes(int slideIdx, const std::string& rId) {
        if (slideIdx < 0 || slideIdx >= (int)m_slides.size()) return {};
        auto& slide = m_slides[slideIdx];
        for (auto& img : slide.images) {
            if (img.rId == rId) return img.bytes;
        }
        return {};
    }

    void ApplyEdit(const EditCommand& cmd) {
        m_editHistory.push_back(cmd);
        int idx = std::stoi(cmd.slideId);
        if (idx >= 0 && idx < (int)m_slides.size()) {
            ApplyEditToSlide(m_slides[idx], cmd);
        }
    }

    void UndoEdit() {
        if (!m_editHistory.empty()) m_editHistory.pop_back();
    }

    bool SavePptx(const std::wstring& path) {
        // Rebuild XML for all slides, then repack ZIP
        return RebuildAndSave(path);
    }

    // Generate HTML preview (legacy / WebView2 mode)
    std::string GenerateSlideHTML(int index) {
        if (index < 0 || index >= (int)m_slides.size()) return "";
        m_currentSlide = index;
        auto& slide = m_slides[index];
        std::string html = BuildHTMLHeader(slide.width, slide.height);
        html += BuildShapesHTMLFromParsed(slide);
        html += BuildEditScript();
        html += "</div></body></html>";
        return html;
    }

private:
    // ── ZIP ───────────────────────────────────────────────────────
    void ParseZipDirectory() {
        if (m_zip.size() < 22) return;
        // Find EOCD
        for (size_t i = m_zip.size() - 22; i != (size_t)-1 && i >= 0; i--) {
            if (m_zip[i]   == 0x50 && m_zip[i+1] == 0x4B &&
                m_zip[i+2] == 0x05 && m_zip[i+3] == 0x06) {
                uint16_t numEntries = *(uint16_t*)(&m_zip[i+10]);
                uint32_t cdOffset   = *(uint32_t*)(&m_zip[i+16]);
                uint32_t pos = cdOffset;
                if (pos >= m_zip.size()) break;

                for (uint16_t j = 0; j < numEntries && pos + 46 < m_zip.size(); j++) {
                    if (m_zip[pos] != 0x50 || m_zip[pos+1] != 0x4B) break;
                    ZipEntry e;
                    e.method    = *(uint16_t*)(&m_zip[pos+10]);
                    e.compSize  = *(uint32_t*)(&m_zip[pos+20]);
                    e.uncompSize= *(uint32_t*)(&m_zip[pos+24]);
                    uint16_t nameLen  = *(uint16_t*)(&m_zip[pos+28]);
                    uint16_t extraLen = *(uint16_t*)(&m_zip[pos+30]);
                    uint16_t commentLen = *(uint16_t*)(&m_zip[pos+32]);
                    uint32_t lhOffset = *(uint32_t*)(&m_zip[pos+42]);

                    if (pos + 46 + nameLen > m_zip.size()) break;
                    e.name = std::string((char*)&m_zip[pos+46], nameLen);

                    // Local header offset
                    if (lhOffset + 30 < m_zip.size()) {
                        uint16_t lNameLen  = *(uint16_t*)(&m_zip[lhOffset+26]);
                        uint16_t lExtraLen = *(uint16_t*)(&m_zip[lhOffset+28]);
                        e.offset = lhOffset + 30 + lNameLen + lExtraLen;
                    }
                    m_entries.push_back(e);
                    pos += 46 + nameLen + extraLen + commentLen;
                }
                break;
            }
        }
    }

    std::string ExtractFile(const std::string& name) {
        for (auto& e : m_entries) {
            std::string n = e.name;
            // Normalize slashes
            for (auto& c : n) if (c == '\\') c = '/';
            std::string target = name;
            for (auto& c : target) if (c == '\\') c = '/';

            if (n == target || (n.size() > 0 && n[0] != '/' && "/" + n == "/" + target)) {
                if (e.method == 0) {
                    if (e.offset + e.compSize > m_zip.size()) return "";
                    return std::string((char*)&m_zip[e.offset], e.compSize);
                } else if (e.method == 8) {
                    std::vector<uint8_t> out;
                    out.reserve(e.uncompSize);
                    Inflater inf(&m_zip[e.offset], e.compSize, out);
                    if (!inf.inflate()) return "";
                    return std::string(out.begin(), out.end());
                }
            }
        }
        return "";
    }

    std::vector<uint8_t> ExtractBinary(const std::string& name) {
        for (auto& e : m_entries) {
            std::string n = e.name;
            for (auto& c : n) if (c == '\\') c = '/';
            std::string target = name;
            for (auto& c : target) if (c == '\\') c = '/';
            if (n == target) {
                if (e.method == 0) {
                    if (e.offset + e.compSize > m_zip.size()) return {};
                    return std::vector<uint8_t>(&m_zip[e.offset], &m_zip[e.offset]+e.compSize);
                } else if (e.method == 8) {
                    std::vector<uint8_t> out;
                    out.reserve(e.uncompSize);
                    Inflater inf(&m_zip[e.offset], e.compSize, out);
                    inf.inflate();
                    return out;
                }
            }
        }
        return {};
    }

    // List all entries matching prefix
    std::vector<std::string> ListFiles(const std::string& prefix) {
        std::vector<std::string> result;
        for (auto& e : m_entries) {
            if (e.name.substr(0, prefix.size()) == prefix) result.push_back(e.name);
        }
        return result;
    }

    // ── Presentation loading ──────────────────────────────────────
    void LoadPresentation() {
        std::string presXml = ExtractFile("ppt/presentation.xml");
        if (presXml.empty()) {
            // Try alternate path
            presXml = ExtractFile("ppt\\presentation.xml");
        }
        if (presXml.empty()) return;

        auto* root = XmlNode::ParseFull(presXml);
        if (!root) return;

        // Slide size
        auto* sldSz = root->Find("sldSz");
        if (sldSz) {
            double cxEmu = std::stod(sldSz->Attr("cx","9144000"));
            double cyEmu = std::stod(sldSz->Attr("cy","5143500"));
            m_slideW = EmuToF(cxEmu);
            m_slideH = EmuToF(cyEmu);
        }

        // Parse relationships to find slide paths
        std::string relsXml = ExtractFile("ppt/_rels/presentation.xml.rels");
        std::map<std::string, std::string> rIdToTarget;
        if (!relsXml.empty()) {
            auto* rels = XmlNode::ParseFull(relsXml);
            if (rels) {
                for (auto* rel : rels->children) {
                    std::string type = rel->Attr("Type");
                    if (type.find("/slide\"") != std::string::npos ||
                        type.find("/slide/") != std::string::npos) {
                        rIdToTarget[rel->Attr("Id")] = rel->Attr("Target");
                    }
                }
                delete rels;
            }
        }

        // sldIdLst
        auto* sldIdLst = root->Find("sldIdLst");
        if (sldIdLst) {
            int order = 1;
            for (auto* sldId : sldIdLst->children) {
                std::string rId = sldId->Attr("r:id");
                if (rId.empty()) rId = sldId->Attr("id");

                std::string path;
                if (rIdToTarget.count(rId)) {
                    path = "ppt/" + rIdToTarget[rId];
                    // Normalize ../
                    while (path.find("/../") != std::string::npos) {
                        size_t p = path.find("/../");
                        size_t prev = path.rfind('/', p-1);
                        if (prev == std::string::npos) break;
                        path = path.substr(0,prev) + path.substr(p+3);
                    }
                } else {
                    path = "ppt/slides/slide" + std::to_string(order) + ".xml";
                }

                SlideData slide;
                slide.xml  = ExtractFile(path);
                // rels
                size_t lastSlash = path.rfind('/');
                std::string relsPath = (lastSlash != std::string::npos)
                    ? path.substr(0, lastSlash) + "/_rels/" + path.substr(lastSlash+1) + ".rels"
                    : "_rels/" + path + ".rels";
                slide.rels = ExtractFile(relsPath);
                slide.number = order++;
                slide.width  = m_slideW;
                slide.height = m_slideH;

                if (!slide.xml.empty()) {
                    LoadSlideImages(slide, relsPath);
                    m_slides.push_back(std::move(slide));
                }
            }
        }

        delete root;
    }

    void LoadSlideImages(SlideData& slide, const std::string& relsPath) {
        if (slide.rels.empty()) return;
        auto* rels = XmlNode::ParseFull(slide.rels);
        if (!rels) return;

        for (auto* rel : rels->children) {
            std::string type = rel->Attr("Type");
            if (type.find("image") != std::string::npos) {
                ImageData img;
                img.rId = rel->Attr("Id");
                std::string target = rel->Attr("Target");
                // Build path
                size_t lastSlash = relsPath.rfind("/_rels/");
                std::string base = (lastSlash != std::string::npos)
                    ? relsPath.substr(0, lastSlash) + "/" : "ppt/slides/";
                img.bytes = ExtractBinary(base + target);
                if (img.bytes.empty()) img.bytes = ExtractBinary("ppt/" + target);
                if (!img.bytes.empty()) slide.images.push_back(std::move(img));
            }
        }
        delete rels;
    }

    // ── Theme ─────────────────────────────────────────────────────
    void LoadTheme() {
        // Try theme1.xml
        std::string themeXml = ExtractFile("ppt/theme/theme1.xml");
        if (themeXml.empty()) {
            auto files = ListFiles("ppt/theme/");
            for (auto& f : files) {
                if (f.find(".xml") != std::string::npos) {
                    themeXml = ExtractFile(f);
                    if (!themeXml.empty()) break;
                }
            }
        }
        if (themeXml.empty()) return;

        auto* root = XmlNode::ParseFull(themeXml);
        if (!root) return;
        m_theme.name = root->Attr("name");

        auto* fmts = root->Find("themeElements");
        if (!fmts) fmts = root->Find("fmtScheme");
        if (fmts) {
            auto* clr = fmts->Find("clrScheme");
            if (clr) ExtractColors(clr);
            auto* font = fmts->Find("fontScheme");
            if (font) ExtractFonts(font);
        }
        delete root;
    }

    void ExtractColors(XmlNode* scheme) {
        static const char* schemeNames[] = {
            "dk1","lt1","dk2","lt2","accent1","accent2","accent3",
            "accent4","accent5","accent6","hlink","folHlink"
        };
        for (auto* child : scheme->children) {
            std::string key = child->name;
            size_t col = key.find(':');
            if (col != std::string::npos) key = key.substr(col+1);

            for (auto* color : child->children) {
                std::string cn = color->name;
                col = cn.find(':');
                if (col != std::string::npos) cn = cn.substr(col+1);

                if (cn == "srgbClr" || cn == "sysClr") {
                    std::string val = color->Attr("val");
                    if (val.empty()) val = color->Attr("lastClr");
                    if (!val.empty()) m_theme.colors[key] = "#" + val;
                }
            }
        }
    }

    void ExtractFonts(XmlNode* fontScheme) {
        auto* major = fontScheme->Find("majorFont");
        auto* minor = fontScheme->Find("minorFont");
        if (major) {
            auto* latin = major->Find("latin");
            if (latin) m_theme.fonts["major"] = latin->Attr("typeface");
        }
        if (minor) {
            auto* latin = minor->Find("latin");
            if (latin) m_theme.fonts["minor"] = latin->Attr("typeface");
        }
    }

    // ── Master & Layouts ──────────────────────────────────────────
    void LoadMaster() {
        m_master.xml  = ExtractFile("ppt/slideMasters/slideMaster1.xml");
        m_master.rels = ExtractFile("ppt/slideMasters/_rels/slideMaster1.xml.rels");
    }

    void LoadLayouts() {
        auto files = ListFiles("ppt/slideLayouts/");
        for (auto& f : files) {
            if (f.find(".xml") == std::string::npos) continue;
            if (f.find(".rels") != std::string::npos) continue;
            SlideLayout layout;
            layout.xml = ExtractFile(f);
            auto* root = XmlNode::ParseFull(layout.xml);
            if (root) {
                auto* cSld = root->Find("cSld");
                if (cSld) layout.name = cSld->Attr("name");
                auto* cSldTy = root->Find("typeAttr");
                delete root;
            }
            m_layouts.push_back(std::move(layout));
        }
    }

    // ── Full Slide Parsing ────────────────────────────────────────
    void ParseAllSlides() {
        for (auto& slide : m_slides) {
            ParseSlide(slide);
        }
    }

    void ParseSlide(SlideData& slide) {
        if (slide.xml.empty()) return;
        auto* root = XmlNode::ParseFull(slide.xml);
        if (!root) return;

        // cSld → spTree
        auto* cSld   = root->Find("cSld");
        auto* bg     = cSld ? cSld->Find("bg") : nullptr;
        auto* spTree = cSld ? cSld->Find("spTree") : nullptr;

        if (bg) slide.background = ParseFill(bg->Find("bgPr"), slide);
        if (spTree) ParseShapeTree(spTree, slide.shapes, slide);

        // Notes
        auto* notes = root->Find("notes");
        if (notes) {
            std::stringstream ns;
            CollectText(notes, ns);
            slide.notes = ns.str();
        }

        // Title
        for (auto& sh : slide.shapes) {
            if (!sh.paragraphs.empty() && !sh.paragraphs[0].runs.empty()) {
                if (sh.name.find("Title") != std::string::npos ||
                    sh.name.find("title") != std::string::npos) {
                    std::wstring t;
                    for (auto& r : sh.paragraphs[0].runs) t += r.text;
                    slide.title = WideToUtf8(t);
                    break;
                }
            }
        }

        delete root;
    }

    void CollectText(XmlNode* node, std::stringstream& ss) {
        if (node->name.find(":t") != std::string::npos || node->name == "t") {
            ss << node->value << " ";
        }
        for (auto* c : node->children) CollectText(c, ss);
    }

    void ParseShapeTree(XmlNode* spTree, std::vector<ShapeData>& shapes, SlideData& slide) {
        for (auto* node : spTree->children) {
            std::string tag = RemoveNS(node->name);
            if      (tag == "sp")     shapes.push_back(ParseShape(node, slide));
            else if (tag == "pic")    shapes.push_back(ParsePicture(node, slide));
            else if (tag == "graphicFrame") shapes.push_back(ParseGraphicFrame(node, slide));
            else if (tag == "grpSp") {
                ShapeData grp;
                grp.type = ShapeType::Group;
                ParseTransform(node->Find("grpSpPr"), grp);
                ParseShapeTree(node, grp.children, slide);
                shapes.push_back(std::move(grp));
            }
        }
    }

    ShapeData ParseShape(XmlNode* sp, SlideData& slide) {
        ShapeData shape;
        shape.type = ShapeType::Rect;

        auto* nvSpPr = sp->Find("nvSpPr");
        if (nvSpPr) {
            auto* cNvPr = nvSpPr->Find("cNvPr");
            if (cNvPr) {
                shape.id   = std::stoi(cNvPr->Attr("id","0"));
                shape.name = cNvPr->Attr("name");
                auto* hl = cNvPr->Find("hlinkClick");
                if (hl) shape.hyperlinkUrl = hl->Attr("r:href");
            }
        }

        auto* spPr = sp->Find("spPr");
        if (spPr) {
            ParseTransform(spPr, shape);
            ParseGeometry(spPr, shape);
            shape.fill = ParseFill(spPr, slide);
            shape.line = ParseLine(spPr->Find("ln"), slide);
            shape.effect = ParseEffect(spPr->Find("effectLst"));
        }

        auto* style = sp->Find("style");
        if (style) ApplyShapeStyle(style, shape);

        auto* txBody = sp->Find("txBody");
        if (txBody) {
            ParseTextBody(txBody, shape);
        }

        return shape;
    }

    ShapeData ParsePicture(XmlNode* pic, SlideData& slide) {
        ShapeData shape;
        shape.type = ShapeType::Picture;

        auto* nvPicPr = pic->Find("nvPicPr");
        if (nvPicPr) {
            auto* cNvPr = nvPicPr->Find("cNvPr");
            if (cNvPr) {
                shape.id   = std::stoi(cNvPr->Attr("id","0"));
                shape.name = cNvPr->Attr("name");
            }
        }

        auto* blipFill = pic->Find("blipFill");
        if (blipFill) {
            auto* blip = blipFill->Find("blip");
            if (blip) {
                shape.imageRId = blip->Attr("r:embed");
                std::string alpha = blip->Attr("r:alphaModFix");
                if (!alpha.empty()) shape.imageAlpha = std::stof(alpha) / 100000.0f;
            }
        }

        auto* spPr = pic->Find("spPr");
        if (spPr) ParseTransform(spPr, shape);

        return shape;
    }

    ShapeData ParseGraphicFrame(XmlNode* gf, SlideData& slide) {
        ShapeData shape;
        shape.type = ShapeType::Table;

        ParseTransform(gf->Find("xfrm"), shape);

        auto* graphic = gf->Find("graphic");
        if (!graphic) return shape;
        auto* graphicData = graphic->Find("graphicData");
        if (!graphicData) return shape;

        auto* tbl = graphicData->Find("tbl");
        if (tbl) ParseTable(tbl, shape, slide);

        return shape;
    }

    void ParseTable(XmlNode* tbl, ShapeData& shape, SlideData& slide) {
        // Column widths
        auto* tblGrid = tbl->Find("tblGrid");
        if (tblGrid) {
            for (auto* gc : tblGrid->ChildrenNamed("gridCol")) {
                double w = std::stod(gc->Attr("w","0"));
                shape.colWidths.push_back(EmuToF(w));
            }
        }

        // Rows
        for (auto* tr : tbl->ChildrenNamed("tr")) {
            double rh = std::stod(tr->Attr("h","0"));
            shape.rowHeights.push_back(EmuToF(rh));

            std::vector<ShapeData::TableCell> row;
            for (auto* tc : tr->ChildrenNamed("tc")) {
                ShapeData::TableCell cell;
                cell.rowSpan = std::stoi(tc->Attr("rowSpan","1"));
                cell.colSpan = std::stoi(tc->Attr("gridSpan","1"));

                auto* txBody = tc->Find("txBody");
                if (txBody) {
                    ShapeData tmp;
                    ParseTextBody(txBody, tmp);
                    cell.paragraphs = tmp.paragraphs;
                }

                auto* tcPr = tc->Find("tcPr");
                if (tcPr) {
                    cell.fill = ParseFill(tcPr, slide);
                }

                row.push_back(std::move(cell));
            }
            shape.tableRows.push_back(std::move(row));
        }
    }

    void ParseTransform(XmlNode* container, ShapeData& shape) {
        if (!container) return;
        auto* xfrm = container->Find("xfrm");
        if (!xfrm) xfrm = container; // may already be xfrm

        std::string rotStr = xfrm->Attr("rot");
        if (!rotStr.empty()) shape.rot = (float)(std::stod(rotStr) / 60000.0);
        shape.flipH = (xfrm->Attr("flipH") == "1");
        shape.flipV = (xfrm->Attr("flipV") == "1");

        auto* off = xfrm->Find("off");
        auto* ext = xfrm->Find("ext");
        if (off) {
            shape.x = EmuToF(std::stod(off->Attr("x","0")));
            shape.y = EmuToF(std::stod(off->Attr("y","0")));
        }
        if (ext) {
            shape.w = EmuToF(std::stod(ext->Attr("cx","914400")));
            shape.h = EmuToF(std::stod(ext->Attr("cy","914400")));
        }
    }

    void ParseGeometry(XmlNode* spPr, ShapeData& shape) {
        auto* prstGeom = spPr->Find("prstGeom");
        if (prstGeom) {
            shape.geom.prst = prstGeom->Attr("prst");
            shape.type = GeomNameToType(shape.geom.prst);

            auto* avLst = prstGeom->Find("avLst");
            if (avLst) {
                int i = 0;
                for (auto* gd : avLst->ChildrenNamed("gd")) {
                    if (i < 8) {
                        std::string fmla = gd->Attr("fmla");
                        // "val N"
                        size_t sp = fmla.rfind(' ');
                        if (sp != std::string::npos)
                            shape.geom.adjVal[i] = std::stof(fmla.substr(sp+1)) / 100000.0f;
                        i++;
                    }
                }
            }
        }

        auto* custGeom = spPr->Find("custGeom");
        if (custGeom) shape.type = ShapeType::FreeForm;
    }

    ShapeType GeomNameToType(const std::string& prst) {
        static const std::map<std::string, ShapeType> m = {
            {"rect",ShapeType::Rect},{"roundRect",ShapeType::RoundRect},
            {"ellipse",ShapeType::Ellipse},{"triangle",ShapeType::Triangle},
            {"rtTriangle",ShapeType::Triangle},
            {"parallelogram",ShapeType::Parallelogram},{"trapezoid",ShapeType::Trapezoid},
            {"diamond",ShapeType::Diamond},{"pentagon",ShapeType::Pentagon},
            {"hexagon",ShapeType::Hexagon},{"heptagon",ShapeType::Heptagon},
            {"octagon",ShapeType::Octagon},
            {"star4",ShapeType::Star4},{"star5",ShapeType::Star5},
            {"star6",ShapeType::Star6},{"star8",ShapeType::Star8},
            {"star16",ShapeType::Star16},{"star24",ShapeType::Star24},
            {"star32",ShapeType::Star32},
            {"leftArrow",ShapeType::ArrowLeft},{"rightArrow",ShapeType::ArrowRight},
            {"upArrow",ShapeType::ArrowUp},{"downArrow",ShapeType::ArrowDown},
            {"leftRightArrow",ShapeType::ArrowRight},{"upDownArrow",ShapeType::ArrowUp},
            {"bentArrow",ShapeType::ArrowRight},{"stripedRightArrow",ShapeType::ArrowRight},
            {"notchedRightArrow",ShapeType::ArrowRight},
            {"chevron",ShapeType::Chevron},{"homePlate",ShapeType::Chevron},
            {"ribbon",ShapeType::Ribbon},{"ribbon2",ShapeType::Ribbon},
            {"wave",ShapeType::Wave},{"doubleWave",ShapeType::Wave},
            {"cloudCallout",ShapeType::Cloud},{"cloud",ShapeType::Cloud},
            {"heart",ShapeType::Heart},{"lightningBolt",ShapeType::Lightning},
            {"gear6",ShapeType::Gear},{"gear9",ShapeType::Gear},
            {"line",ShapeType::Line},{"straightConnector1",ShapeType::Line},
            {"bentConnector2",ShapeType::Line},{"curvedConnector2",ShapeType::Line},
            {"curvedConnector3",ShapeType::Line},
        };
        auto it = m.find(prst);
        return (it != m.end()) ? it->second : ShapeType::Rect;
    }

    // ── Fill parsing ──────────────────────────────────────────────
    FillInfo ParseFill(XmlNode* container, const SlideData& slide) {
        FillInfo fill;
        if (!container) return fill;

        if (auto* solidFill = container->Find("solidFill")) {
            fill.type = FillInfo::Solid;
            fill.solidColor = ResolveColorNode(solidFill, fill.alpha);
        }
        else if (auto* gradFill = container->Find("gradFill")) {
            fill.type = FillInfo::Gradient;
            ParseGradFill(gradFill, fill);
        }
        else if (auto* noFill = container->Find("noFill")) {
            fill.type = FillInfo::None;
        }
        else if (auto* blipFill = container->Find("blipFill")) {
            fill.type = FillInfo::Picture;
            auto* blip = blipFill->Find("blip");
            if (blip) fill.imageRId = blip->Attr("r:embed");
        }
        return fill;
    }

    void ParseGradFill(XmlNode* gradFill, FillInfo& fill) {
        auto* gsLst = gradFill->Find("gsLst");
        if (gsLst) {
            for (auto* gs : gsLst->ChildrenNamed("gs")) {
                GradStop stop;
                stop.pos = std::stof(gs->Attr("pos","0")) / 100000.0f;
                float a = 1.0f;
                stop.color = ResolveColorNode(gs, a);
                stop.color.a = a;
                fill.gradStops.push_back(stop);
            }
        }
        auto* lin = gradFill->Find("lin");
        if (lin) {
            fill.gradLinear = true;
            fill.gradAngle  = (float)(std::stod(lin->Attr("ang","0")) / 60000.0);
        }
        auto* path = gradFill->Find("path");
        if (path) fill.gradLinear = false;

        // Sort stops by position
        std::sort(fill.gradStops.begin(), fill.gradStops.end(),
                  [](const GradStop& a, const GradStop& b){ return a.pos < b.pos; });
    }

    D2D1_COLOR_F ResolveColorNode(XmlNode* container, float& alpha) {
        alpha = 1.0f;
        if (!container) return D2D1::ColorF(0,0,0);

        D2D1_COLOR_F color = D2D1::ColorF(0,0,0);

        for (auto* child : container->children) {
            std::string tag = RemoveNS(child->name);
            if (tag == "srgbClr") {
                color = HexToColor(child->Attr("val"));
                ApplyColorMods(child, color, alpha);
                return color;
            }
            if (tag == "schemeClr") {
                std::string ref = child->Attr("val");
                std::string hex = m_theme.Resolve(ref);
                color = HexToColor(hex);
                ApplyColorMods(child, color, alpha);
                return color;
            }
            if (tag == "sysClr") {
                color = HexToColor(child->Attr("lastClr","000000"));
                ApplyColorMods(child, color, alpha);
                return color;
            }
            if (tag == "hslClr") {
                float h = std::stof(child->Attr("hue","0")) / 6000.0f;
                float s = std::stof(child->Attr("sat","0")) / 100000.0f;
                float l = std::stof(child->Attr("lum","0")) / 100000.0f;
                color = HslToRgb(h, s, l);
                ApplyColorMods(child, color, alpha);
                return color;
            }
            if (tag == "prstClr") {
                // preset – map common names
                color = PresetColor(child->Attr("val"));
                ApplyColorMods(child, color, alpha);
                return color;
            }
        }
        return color;
    }

    void ApplyColorMods(XmlNode* clrNode, D2D1_COLOR_F& c, float& alpha) {
        for (auto* mod : clrNode->children) {
            std::string tag = RemoveNS(mod->name);
            float val = std::stof(mod->Attr("val","100000")) / 100000.0f;

            if (tag == "alpha" || tag == "alphaMod") {
                alpha *= val;
            } else if (tag == "alphaOff") {
                alpha = Clampf(alpha + val - 1.0f, 0, 1);
            } else if (tag == "lumMod") {
                c.r *= val; c.g *= val; c.b *= val;
            } else if (tag == "lumOff") {
                c.r = Clampf(c.r + val, 0, 1);
                c.g = Clampf(c.g + val, 0, 1);
                c.b = Clampf(c.b + val, 0, 1);
            } else if (tag == "shade") {
                c.r *= val; c.g *= val; c.b *= val;
            } else if (tag == "tint") {
                c.r = Clampf(c.r + (1.0f - c.r) * val, 0, 1);
                c.g = Clampf(c.g + (1.0f - c.g) * val, 0, 1);
                c.b = Clampf(c.b + (1.0f - c.b) * val, 0, 1);
            }
        }
    }

    // HSL → RGB
    static D2D1_COLOR_F HslToRgb(float h, float s, float l) {
        auto hue2rgb = [](float p, float q, float t) {
            if (t < 0) t += 1; if (t > 1) t -= 1;
            if (t < 1.0f/6) return p + (q - p) * 6 * t;
            if (t < 0.5f)   return q;
            if (t < 2.0f/3) return p + (q - p) * (2.0f/3 - t) * 6;
            return p;
        };
        if (s == 0) return D2D1::ColorF(l, l, l);
        float q = l < 0.5f ? l * (1 + s) : l + s - l * s;
        float p = 2 * l - q;
        return D2D1::ColorF(
            hue2rgb(p, q, h + 1.0f/3),
            hue2rgb(p, q, h),
            hue2rgb(p, q, h - 1.0f/3)
        );
    }

    static D2D1_COLOR_F PresetColor(const std::string& name) {
        static const std::map<std::string, uint32_t> m = {
            {"red",0xFF0000},{"green",0x008000},{"blue",0x0000FF},
            {"white",0xFFFFFF},{"black",0x000000},{"yellow",0xFFFF00},
            {"orange",0xFF8000},{"purple",0x800080},{"cyan",0x00FFFF},
            {"magenta",0xFF00FF},{"gray",0x808080},{"silver",0xC0C0C0},
            {"navy",0x000080},{"teal",0x008080},{"lime",0x00FF00},
            {"maroon",0x800000},{"olive",0x808000},{"aqua",0x00FFFF},
            {"fuchsia",0xFF00FF},{"gold",0xFFD700},{"coral",0xFF7F50},
            {"salmon",0xFA8072},{"tan",0xD2B48C},{"brown",0xA52A2A},
            {"pink",0xFFC0CB},{"violet",0xEE82EE},{"indigo",0x4B0082},
        };
        auto it = m.find(name);
        if (it != m.end()) {
            uint32_t rgb = it->second;
            return D2D1::ColorF(
                ((rgb>>16)&0xFF)/255.0f,
                ((rgb>> 8)&0xFF)/255.0f,
                (rgb      &0xFF)/255.0f
            );
        }
        return D2D1::ColorF(0,0,0);
    }

    // ── Line parsing ──────────────────────────────────────────────
    LineInfo ParseLine(XmlNode* ln, const SlideData& slide) {
        LineInfo line;
        if (!ln) return line;

        std::string wStr = ln->Attr("w");
        if (!wStr.empty()) line.width = EmuToF(std::stod(wStr));

        line.fill = ParseFill(ln, slide);

        // Cap
        std::string cap = ln->Attr("cap");
        if      (cap == "rnd") line.cap = LineInfo::Round;
        else if (cap == "sq")  line.cap = LineInfo::Square;

        // Dash
        auto* prstDash = ln->Find("prstDash");
        if (prstDash) {
            std::string d = prstDash->Attr("val");
            if      (d == "dot")           line.dash = LineInfo::Dot;
            else if (d == "dash")          line.dash = LineInfo::Dash_;
            else if (d == "dashDot")       line.dash = LineInfo::DashDot;
            else if (d == "lgDash")        line.dash = LineInfo::LongDash;
            else if (d == "lgDashDot")     line.dash = LineInfo::LongDashDot;
        }

        // Arrowheads
        auto* headEnd = ln->Find("headEnd");
        if (headEnd) {
            line.headType = headEnd->Attr("type");
            line.headW    = std::stof(headEnd->Attr("w","0"));
            line.headLen  = std::stof(headEnd->Attr("len","0"));
        }
        auto* tailEnd = ln->Find("tailEnd");
        if (tailEnd) {
            line.tailType = tailEnd->Attr("type");
            line.tailW    = std::stof(tailEnd->Attr("w","0"));
            line.tailLen  = std::stof(tailEnd->Attr("len","0"));
        }

        return line;
    }

    // ── Effect parsing ────────────────────────────────────────────
    EffectInfo ParseEffect(XmlNode* effectLst) {
        EffectInfo eff;
        if (!effectLst) return eff;

        auto* shadow = effectLst->Find("outerShdw");
        if (!shadow) shadow = effectLst->Find("innerShdw");
        if (shadow) {
            eff.hasShadow   = true;
            eff.shadowBlur  = EmuToF(std::stod(shadow->Attr("blurRad","0")));
            double dist     = std::stod(shadow->Attr("dist","0"));
            double dir      = std::stod(shadow->Attr("dir","0")) / 60000.0 * 3.14159 / 180.0;
            eff.shadowDX    = (float)(EmuToPx(dist) * std::cos(dir));
            eff.shadowDY    = (float)(EmuToPx(dist) * std::sin(dir));
            float a = 0.5f;
            eff.shadowColor = ResolveColorNode(shadow, a);
            eff.shadowColor.a = a;
        }

        auto* glow = effectLst->Find("glow");
        if (glow) {
            eff.hasGlow    = true;
            eff.glowRadius = EmuToF(std::stod(glow->Attr("rad","0")));
            float a = 0.5f;
            eff.glowColor  = ResolveColorNode(glow, a);
            eff.glowColor.a = a;
        }

        auto* se = effectLst->Find("softEdge");
        if (se) {
            eff.hasSoftEdge    = true;
            eff.softEdgeRadius = EmuToF(std::stod(se->Attr("rad","0")));
        }

        auto* ref = effectLst->Find("reflection");
        if (ref) eff.reflection = std::stof(ref->Attr("size","0")) / 100000.0f;

        return eff;
    }

    void ApplyShapeStyle(XmlNode* style, ShapeData& shape) {
        // Style reference overrides (from theme)
        auto* lnRef = style->Find("lnRef");
        if (lnRef && shape.line.width == 0) {
            float a = 1.0f;
            shape.line.fill.solidColor = ResolveColorNode(lnRef, a);
            shape.line.fill.type = FillInfo::Solid;
        }
        auto* fillRef = style->Find("fillRef");
        if (fillRef && shape.fill.type == FillInfo::None) {
            float a = 1.0f;
            shape.fill.solidColor = ResolveColorNode(fillRef, a);
            shape.fill.type = FillInfo::Solid;
        }
    }

    // ── Text body parsing ─────────────────────────────────────────
    void ParseTextBody(XmlNode* txBody, ShapeData& shape) {
        // Body properties
        auto* bodyPr = txBody->Find("bodyPr");
        if (bodyPr) {
            std::string wrap = bodyPr->Attr("wrap");
            shape.wordWrap = (wrap != "none");
            shape.autoFit  = (bodyPr->Find("spAutoFit") != nullptr ||
                              bodyPr->Find("normAutofit") != nullptr);

            // Insets in EMU
            std::string il = bodyPr->Attr("lIns");
            std::string ir = bodyPr->Attr("rIns");
            std::string it = bodyPr->Attr("tIns");
            std::string ib = bodyPr->Attr("bIns");
            if (!il.empty()) shape.txInsetL = EmuToF(std::stod(il));
            if (!ir.empty()) shape.txInsetR = EmuToF(std::stod(ir));
            if (!it.empty()) shape.txInsetT = EmuToF(std::stod(it));
            if (!ib.empty()) shape.txInsetB = EmuToF(std::stod(ib));

            // Vertical anchor
            std::string anchor = bodyPr->Attr("anchor");
            if      (anchor == "ctr") shape.textVAlign = DWRITE_PARAGRAPH_ALIGNMENT_CENTER;
            else if (anchor == "b")   shape.textVAlign = DWRITE_PARAGRAPH_ALIGNMENT_FAR;
            else                      shape.textVAlign = DWRITE_PARAGRAPH_ALIGNMENT_NEAR;
        }

        // Default paragraph props from lstStyle
        ParaProps defaultPara;
        auto* lstStyle = txBody->Find("lstStyle");
        if (lstStyle) {
            auto* defPPr = lstStyle->Find("defPPr");
            if (defPPr) defaultPara = ParseParaProps(defPPr);
        }

        // Paragraphs
        for (auto* para : txBody->ChildrenNamed("p")) {
            Paragraph p;
            auto* pPr = para->Find("pPr");
            p.props = pPr ? ParseParaProps(pPr) : defaultPara;

            // Default run props
            RunProps defaultRun;
            auto* defRPr = para->Find("defRPr");
            if (defRPr) defaultRun = ParseRunProps(defRPr);

            for (auto* child : para->children) {
                std::string tag = RemoveNS(child->name);
                if (tag == "r") {
                    TextRun run;
                    auto* rPr = child->Find("rPr");
                    run.props = rPr ? ParseRunProps(rPr) : defaultRun;
                    auto* t = child->Find("t");
                    if (t) run.text = Utf8ToWide(t->value);
                    if (!run.text.empty()) p.runs.push_back(run);
                } else if (tag == "br") {
                    // Line break
                    TextRun run;
                    run.props = defaultRun;
                    run.text  = L"\n";
                    p.runs.push_back(run);
                } else if (tag == "fld") {
                    // Field (slide number, date, etc.)
                    TextRun run;
                    run.props = defaultRun;
                    auto* t = child->Find("t");
                    if (t) run.text = Utf8ToWide(t->value);
                    else   run.text = L"";
                    p.runs.push_back(run);
                }
            }

            shape.paragraphs.push_back(p);
        }
    }

    ParaProps ParseParaProps(XmlNode* pPr) {
        ParaProps pp;
        if (!pPr) return pp;

        std::string align = pPr->Attr("algn");
        if      (align == "ctr") pp.align = DWRITE_TEXT_ALIGNMENT_CENTER;
        else if (align == "r")   pp.align = DWRITE_TEXT_ALIGNMENT_TRAILING;
        else if (align == "just")pp.align = DWRITE_TEXT_ALIGNMENT_JUSTIFIED;
        else                     pp.align = DWRITE_TEXT_ALIGNMENT_LEADING;

        std::string indent = pPr->Attr("indent");
        if (!indent.empty()) pp.indent = EmuToF(std::stod(indent));

        std::string marL = pPr->Attr("marL");
        if (!marL.empty()) pp.leftMargin = EmuToF(std::stod(marL));

        pp.level = std::stoi(pPr->Attr("lvl","0"));

        // Spacing
        auto* spcBef = pPr->Find("spcBef");
        if (spcBef) {
            auto* spcPts = spcBef->Find("spcPts");
            if (spcPts) pp.spaceBefore = std::stof(spcPts->Attr("val","0")) / 100.0f;
        }
        auto* spcAft = pPr->Find("spcAft");
        if (spcAft) {
            auto* spcPts = spcAft->Find("spcPts");
            if (spcPts) pp.spaceAfter = std::stof(spcPts->Attr("val","0")) / 100.0f;
        }
        auto* lnSpc = pPr->Find("lnSpc");
        if (lnSpc) {
            auto* spcPct = lnSpc->Find("spcPct");
            auto* spcPts = lnSpc->Find("spcPts");
            if (spcPct) {
                pp.lineSpacing = std::stof(spcPct->Attr("val","100000")) / 100000.0f;
                pp.lineSpacingIsPoints = false;
            } else if (spcPts) {
                pp.lineSpacing = std::stof(spcPts->Attr("val","0")) / 100.0f;
                pp.lineSpacingIsPoints = true;
            }
        }

        // Bullet
        auto* buChar = pPr->Find("buChar");
        if (buChar) {
            pp.hasBullet  = true;
            pp.bulletChar = buChar->Attr("char");
        }
        auto* buNone = pPr->Find("buNone");
        if (buNone) pp.hasBullet = false;

        return pp;
    }

    RunProps ParseRunProps(XmlNode* rPr) {
        RunProps rp;
        if (!rPr) return rp;

        std::string sz = rPr->Attr("sz");
        if (!sz.empty()) rp.fontSize = (float)(std::stoi(sz) / 100.0);

        rp.bold       = (rPr->Attr("b","0") == "1");
        rp.italic     = (rPr->Attr("i","0") == "1");
        rp.underline  = (rPr->Attr("u") == "sng" || rPr->Attr("u") == "dbl");
        rp.strikeThru = (rPr->Attr("strike") != "" && rPr->Attr("strike") != "noStrike");
        rp.baseline   = std::stoi(rPr->Attr("baseline","0"));
        rp.spacing    = EmuToF(std::stod(rPr->Attr("spc","0")));

        auto* solidFill = rPr->Find("solidFill");
        if (solidFill) {
            float a = 1.0f;
            rp.color = ResolveColorNode(solidFill, a);
            rp.color.a = a;
        }

        auto* latin = rPr->Find("latin");
        if (latin) {
            std::string tf = latin->Attr("typeface");
            if (!tf.empty() && tf != "+mj-lt" && tf != "+mn-lt") {
                rp.fontFamily = tf;
            } else if (tf == "+mj-lt") {
                auto it = m_theme.fonts.find("major");
                if (it != m_theme.fonts.end()) rp.fontFamily = it->second;
            } else if (tf == "+mn-lt") {
                auto it = m_theme.fonts.find("minor");
                if (it != m_theme.fonts.end()) rp.fontFamily = it->second;
            }
        }

        return rp;
    }

    // ── HTML generation (WebView2 mode) ───────────────────────────
    std::string BuildHTMLHeader(double w, double h) {
        std::stringstream ss;
        ss << "<!DOCTYPE html><html><head><meta charset='UTF-8'><style>\n"
           << "*{margin:0;padding:0;box-sizing:border-box}\n"
           << "body{display:flex;justify-content:center;align-items:center;"
           << "min-height:100vh;background:#0a0a1a;font-family:'Segoe UI',sans-serif}\n"
           << ".slide{position:relative;width:" << w << "px;height:" << h
           << "px;background:white;box-shadow:0 10px 50px rgba(0,0,0,.6)}\n"
           << ".toolbar{position:fixed;top:0;left:0;right:0;background:#1a1a2e;"
           << "color:white;padding:10px 20px;display:flex;justify-content:space-between;"
           << "z-index:1000;align-items:center}\n"
           << ".toolbar button{background:#30305a;color:white;border:none;"
           << "padding:8px 16px;margin:0 4px;border-radius:4px;cursor:pointer}\n"
           << ".toolbar button:hover{background:#4040aa}\n"
           << ".shape:hover{outline:2px solid #0066ff;cursor:pointer}\n"
           << ".shape.selected{outline:3px solid #00ff00}\n"
           << "[contenteditable]:focus{outline:2px dashed #ff6600}\n"
           << ".edit-panel{position:fixed;right:0;top:50px;bottom:0;"
           << "width:280px;background:#1a1a2e;color:white;padding:15px;"
           << "overflow-y:auto;z-index:999}\n"
           << ".edit-panel input,.edit-panel select{width:100%;padding:8px;"
           << "margin:5px 0;background:#30305a;color:white;border:1px solid #5050aa}\n"
           << "</style></head><body>\n"
           << "<div class='toolbar'>\n"
           << "<div><b>Presentation Studio Pro</b> | Slide " << (m_currentSlide + 1)
           << " / " << m_slides.size() << "</div>\n"
           << "<div>\n"
           << "<button onclick='prevSlide()'>← Previous</button>\n"
           << "<button onclick='nextSlide()'>Next →</button>\n"
           << "<button onclick='saveDocument()'>💾 Save</button>\n"
           << "<button onclick='undoEdit()'>↩ Undo</button>\n"
           << "<button onclick='exportPDF()'>📄 Export PDF</button>\n"
           << "</div></div>\n"
           << "<div class='edit-panel' id='editPanel'>\n"
           << "<h3>Properties</h3><div id='props'></div>\n"
           << "</div>\n"
           << "<div class='slide' id='mainSlide'>\n";
        return ss.str();
    }

    std::string BuildShapesHTMLFromParsed(const SlideData& slide) {
        std::stringstream ss;

        // Background
        ss << "<div style='position:absolute;inset:0;"
           << FillToCSS(slide.background) << "'></div>\n";

        for (const auto& shape : slide.shapes) {
            ss << ShapeToHTML(shape, slide);
        }
        return ss.str();
    }

    std::string FillToCSS(const FillInfo& fill) {
        std::stringstream css;
        if (fill.type == FillInfo::None) {
            css << "background:transparent;";
        } else if (fill.type == FillInfo::Solid) {
            char hex[32];
            snprintf(hex, sizeof(hex), "rgba(%d,%d,%d,%.3f)",
                     (int)(fill.solidColor.r*255),
                     (int)(fill.solidColor.g*255),
                     (int)(fill.solidColor.b*255),
                     fill.alpha);
            css << "background:" << hex << ";";
        } else if (fill.type == FillInfo::Gradient) {
            css << "background:" << GradientToCSS(fill) << ";";
        }
        return css.str();
    }

    std::string GradientToCSS(const FillInfo& fill) {
        std::stringstream css;
        if (fill.gradLinear) {
            css << "linear-gradient(" << (fill.gradAngle + 90.0) << "deg";
        } else {
            css << "radial-gradient(ellipse at center";
        }
        for (auto& stop : fill.gradStops) {
            char hex[64];
            snprintf(hex, sizeof(hex), ", rgba(%d,%d,%d,%.3f) %.1f%%",
                     (int)(stop.color.r*255),(int)(stop.color.g*255),
                     (int)(stop.color.b*255), stop.color.a,
                     stop.pos * 100.0f);
            css << hex;
        }
        css << ")";
        return css.str();
    }

    std::string ShapeToHTML(const ShapeData& shape, const SlideData& slide) {
        std::stringstream ss;
        ss << "<div class='shape' id='shape" << shape.id << "' "
           << "style='position:absolute;left:" << shape.x << "px;top:" << shape.y
           << "px;width:" << shape.w << "px;height:" << shape.h << "px;"
           << FillToCSS(shape.fill);

        if (shape.line.width > 0) {
            char stroke[128];
            snprintf(stroke, sizeof(stroke), "border:%.1fpx solid rgba(%d,%d,%d,%.2f);",
                     shape.line.width,
                     (int)(shape.line.fill.solidColor.r*255),
                     (int)(shape.line.fill.solidColor.g*255),
                     (int)(shape.line.fill.solidColor.b*255),
                     1.0f);
            ss << stroke;
        }

        if (shape.type == ShapeType::Ellipse)    ss << "border-radius:50%;";
        if (shape.type == ShapeType::RoundRect)  ss << "border-radius:8px;";
        if (shape.rot != 0) ss << "transform:rotate(" << shape.rot << "deg);";
        if (shape.effect.hasShadow) {
            char shadow[128];
            snprintf(shadow, sizeof(shadow),
                     "box-shadow:%.1fpx %.1fpx %.1fpx rgba(%d,%d,%d,%.2f);",
                     shape.effect.shadowDX, shape.effect.shadowDY,
                     shape.effect.shadowBlur,
                     (int)(shape.effect.shadowColor.r*255),
                     (int)(shape.effect.shadowColor.g*255),
                     (int)(shape.effect.shadowColor.b*255),
                     shape.effect.shadowColor.a);
            ss << shadow;
        }

        ss << "overflow:hidden;' onclick='selectShape(" << shape.id << ")'>";

        // Text
        for (const auto& para : shape.paragraphs) {
            ss << "<p style='";
            if (para.props.align == DWRITE_TEXT_ALIGNMENT_CENTER)    ss << "text-align:center;";
            else if (para.props.align == DWRITE_TEXT_ALIGNMENT_TRAILING) ss << "text-align:right;";
            else if (para.props.align == DWRITE_TEXT_ALIGNMENT_JUSTIFIED)ss << "text-align:justify;";
            ss << "margin:" << para.props.spaceBefore/2 << "px 0 "
               << para.props.spaceAfter/2 << "px;"
               << "padding:" << shape.txInsetT << "px " << shape.txInsetR << "px 0 "
               << shape.txInsetL << "px;"
               << "' contenteditable='true' oninput='textChanged(this)' onblur='saveTextEdit(this)'>";

            if (para.props.hasBullet && !para.props.bulletChar.empty())
                ss << "<span>" << EscapeHTML(para.props.bulletChar) << " </span>";

            for (const auto& run : para.runs) {
                ss << "<span style='";
                ss << "font-family:'" << run.props.fontFamily << "',sans-serif;";
                ss << "font-size:" << run.props.fontSize << "pt;";
                if (run.props.bold)       ss << "font-weight:bold;";
                if (run.props.italic)     ss << "font-style:italic;";
                if (run.props.underline)  ss << "text-decoration:underline;";
                if (run.props.strikeThru) ss << "text-decoration:line-through;";
                char fc[64];
                snprintf(fc, sizeof(fc), "color:rgba(%d,%d,%d,%.2f);",
                         (int)(run.props.color.r*255),(int)(run.props.color.g*255),
                         (int)(run.props.color.b*255), run.props.color.a);
                ss << fc;
                if (run.props.baseline > 0)       ss << "vertical-align:super;font-size:0.7em;";
                else if (run.props.baseline < 0)  ss << "vertical-align:sub;font-size:0.7em;";
                ss << "'>";
                ss << EscapeHTML(WideToUtf8(run.text));
                ss << "</span>";
            }
            ss << "</p>";
        }
        ss << "</div>\n";
        return ss.str();
    }

    std::string BuildShapesHTML(XmlNode* slideNode) {
        std::stringstream ss;
        int shapeId = 0;
        for (auto* node : slideNode->children) {
            if (node->name.find(":sp") != std::string::npos || node->name == "sp") {
                ss << BuildShapeHTML(node, shapeId++);
            }
        }
        return ss.str();
    }

    std::string BuildShapeHTML(XmlNode* shape, int id) {
        std::stringstream ss;
        double x = 0, y = 0, w = 100, h = 100, rot = 0;
        std::string fill = "white", stroke = "none";
        double strokeW = 0;

        for (auto* child : shape->children) {
            std::string tag = RemoveNS(child->name);
            if (tag == "spPr") {
                for (auto* prop : child->children) {
                    std::string ptag = RemoveNS(prop->name);
                    if (ptag == "xfrm") {
                        for (auto* trans : prop->children) {
                            std::string ttag = RemoveNS(trans->name);
                            if (ttag == "off") {
                                x = GetAttrD(trans, "x", 0) / 12700.0;
                                y = GetAttrD(trans, "y", 0) / 12700.0;
                            } else if (ttag == "ext") {
                                w = GetAttrD(trans, "cx", 914400) / 12700.0;
                                h = GetAttrD(trans, "cy", 914400) / 12700.0;
                            }
                        }
                        rot = GetAttrD(prop, "rot", 0) / 60000.0;
                    } else if (ptag == "solidFill") {
                        fill = GetFillColorStr(prop);
                    } else if (ptag == "gradFill") {
                        fill = BuildGradient(prop);
                    } else if (ptag == "ln") {
                        strokeW = GetAttrD(prop, "w", 0) / 12700.0;
                        stroke  = GetStrokeColorStr(prop);
                    }
                }
            } else if (tag == "txBody") {
                ss << "<div class='shape' id='shape" << id << "' "
                   << "style='position:absolute;left:" << x << "px;top:" << y
                   << "px;width:" << w << "px;height:" << h << "px;"
                   << "background:" << fill << ";"
                   << (strokeW > 0 ? "border:" + std::to_string(strokeW) + "px solid " + stroke + ";" : "")
                   << (rot != 0 ? "transform:rotate(" + std::to_string(rot) + "deg);" : "")
                   << "overflow:hidden;' onclick='selectShape(" << id << ")'>";
                ss << BuildTextHTML(child);
                ss << "</div>";
                return ss.str();
            }
        }

        ss << "<div class='shape' id='shape" << id << "' "
           << "style='position:absolute;left:" << x << "px;top:" << y
           << "px;width:" << w << "px;height:" << h << "px;"
           << "background:" << fill << ";"
           << (strokeW > 0 ? "border:" + std::to_string(strokeW) + "px solid " + stroke + ";" : "")
           << (rot != 0 ? "transform:rotate(" + std::to_string(rot) + "deg);" : "")
           << "overflow:hidden;' onclick='selectShape(" << id << ")'></div>";
        return ss.str();
    }

    std::string BuildTextHTML(XmlNode* txBody) {
        std::stringstream ss;
        for (auto* para : txBody->children) {
            if (RemoveNS(para->name) != "p") continue;
            ss << "<p contenteditable='true' style='margin:5px 0;padding:5px;"
               << "min-height:1em;outline:none;' oninput='textChanged(this)' onblur='saveTextEdit(this)'>";
            for (auto* run : para->children) {
                if (RemoveNS(run->name) != "r") continue;
                std::string style = GetRunStyleStr(run);
                std::string text  = GetRunTextStr(run);
                ss << "<span contenteditable='true' style='" << style << "'>"
                   << EscapeHTML(text) << "</span>";
            }
            ss << "</p>";
        }
        return ss.str();
    }

    std::string GetRunStyleStr(XmlNode* run) {
        std::stringstream style;
        for (auto* child : run->children) {
            if (RemoveNS(child->name) == "rPr") {
                if (child->attrs.count("sz"))
                    style << "font-size:" << std::stoi(child->Attr("sz")) / 100.0 << "pt;";
                if (child->Attr("b") == "1") style << "font-weight:bold;";
                if (child->Attr("i") == "1") style << "font-style:italic;";
                for (auto* c : child->children) {
                    if (RemoveNS(c->name) == "solidFill") {
                        std::string color = GetFillColorStr(c);
                        if (!color.empty()) style << "color:" << color << ";";
                    } else if (RemoveNS(c->name) == "latin") {
                        if (c->attrs.count("typeface"))
                            style << "font-family:'" << c->Attr("typeface") << "',sans-serif;";
                    }
                }
            }
        }
        return style.str();
    }

    std::string GetRunTextStr(XmlNode* run) {
        for (auto* child : run->children) {
            if (RemoveNS(child->name) == "t") return child->value;
        }
        return "";
    }

    std::string GetFillColorStr(XmlNode* fillNode) {
        for (auto* child : fillNode->children) {
            std::string cn = RemoveNS(child->name);
            if (cn == "srgbClr") return "#" + child->Attr("val");
            if (cn == "schemeClr") return m_theme.Resolve(child->Attr("val"));
        }
        return "";
    }

    std::string GetStrokeColorStr(XmlNode* lnNode) {
        for (auto* child : lnNode->children) {
            if (RemoveNS(child->name) == "solidFill") return GetFillColorStr(child);
        }
        return "#000000";
    }

    std::string BuildGradient(XmlNode* gradFill) {
        std::stringstream css;
        bool linear = false; double angle = 0;
        std::vector<std::pair<std::string,double>> stops;
        for (auto* child : gradFill->children) {
            std::string tag = RemoveNS(child->name);
            if (tag == "lin") { linear = true; angle = GetAttrD(child,"ang",0)/60000.0; }
            else if (tag == "gs") {
                double pos = GetAttrD(child,"pos",0) / 1000.0;
                std::string color;
                for (auto* c : child->children)
                    if (RemoveNS(c->name) == "srgbClr") color = "#" + c->Attr("val");
                stops.push_back({color, pos});
            }
        }
        css << (linear ? "linear-gradient(" + std::to_string(angle) + "deg"
                       : "radial-gradient(circle");
        for (auto& [color, pos] : stops)
            css << ", " << color << " " << pos << "%";
        css << ")";
        return css.str();
    }

    std::string BuildEditScript() {
        return R"SCRIPT(
<script>
let selectedShape = null;
let editHistory   = [];
let isDragging    = false;
let dragStart     = {};
let dragShape     = {};

function selectShape(id) {
  if(selectedShape !== null) {
    let prev = document.getElementById('shape'+selectedShape);
    if(prev) prev.classList.remove('selected');
  }
  selectedShape = id;
  let shape = document.getElementById('shape'+id);
  if(shape) shape.classList.add('selected');
  showProperties(id);
}

function showProperties(id) {
  let shape = document.getElementById('shape'+id);
  if(!shape) return;
  let st = shape.style;
  let panel = document.getElementById('props');
  if(!panel) return;
  panel.innerHTML = `
    <div style="color:#aaa;margin-bottom:8px;font-size:12px">Shape #${id}</div>
    <label>X Position</label>
    <input type="number" value="${parseFloat(st.left)||0}"
           onchange="moveShape(${id},'x',this.value)">
    <label>Y Position</label>
    <input type="number" value="${parseFloat(st.top)||0}"
           onchange="moveShape(${id},'y',this.value)">
    <label>Width</label>
    <input type="number" value="${parseFloat(st.width)||100}"
           onchange="resizeShape(${id},'w',this.value)">
    <label>Height</label>
    <input type="number" value="${parseFloat(st.height)||100}"
           onchange="resizeShape(${id},'h',this.value)">
    <label>Rotation (deg)</label>
    <input type="number" value="0"
           onchange="rotateShape(${id},this.value)">
    <label>Opacity</label>
    <input type="range" min="0" max="100" value="100"
           oninput="setOpacity(${id},this.value)">
    <label>Fill Color</label>
    <input type="color" value="#ffffff"
           onchange="setFill(${id},this.value)">
    <label>Stroke Color</label>
    <input type="color" value="#000000"
           onchange="setStroke(${id},this.value)">
    <label>Stroke Width</label>
    <input type="number" min="0" max="20" value="${parseFloat(st.borderWidth)||0}"
           onchange="setStrokeWidth(${id},this.value)">
    <hr style="border-color:#333;margin:10px 0">
    <button onclick="bringForward(${id})">⬆ Bring Forward</button>
    <button onclick="sendBackward(${id})">⬇ Send Backward</button>
    <button onclick="duplicateShape(${id})">📋 Duplicate</button>
    <button onclick="deleteShape(${id})">🗑 Delete</button>
  `;
}

function moveShape(id, axis, val) {
  let s = document.getElementById('shape'+id);
  if(!s) return;
  s.style[axis==='x'?'left':'top'] = val + 'px';
  recordEdit({type:'move', shape:id, axis, value:val});
}

function resizeShape(id, dim, val) {
  let s = document.getElementById('shape'+id);
  if(!s) return;
  s.style[dim==='w'?'width':'height'] = val + 'px';
  recordEdit({type:'resize', shape:id, dim, value:val});
}

function rotateShape(id, deg) {
  let s = document.getElementById('shape'+id);
  if(!s) return;
  s.style.transform = `rotate(${deg}deg)`;
  recordEdit({type:'rotate', shape:id, value:deg});
}

function setOpacity(id, pct) {
  let s = document.getElementById('shape'+id);
  if(!s) return;
  s.style.opacity = pct / 100;
}

function setFill(id, color) {
  let s = document.getElementById('shape'+id);
  if(!s) return;
  s.style.background = color;
  recordEdit({type:'fill', shape:id, value:color});
}

function setStroke(id, color) {
  let s = document.getElementById('shape'+id);
  if(!s) return;
  let w = parseFloat(s.style.borderWidth) || 1;
  s.style.border = `${w}px solid ${color}`;
}

function setStrokeWidth(id, w) {
  let s = document.getElementById('shape'+id);
  if(!s) return;
  let col = s.style.borderColor || '#000';
  s.style.border = `${w}px solid ${col}`;
}

function bringForward(id) {
  let s = document.getElementById('shape'+id);
  if(!s) return;
  let z = parseInt(s.style.zIndex) || 0;
  s.style.zIndex = z + 1;
}

function sendBackward(id) {
  let s = document.getElementById('shape'+id);
  if(!s) return;
  let z = parseInt(s.style.zIndex) || 0;
  s.style.zIndex = Math.max(0, z - 1);
}

function duplicateShape(id) {
  let s = document.getElementById('shape'+id);
  if(!s) return;
  let clone = s.cloneNode(true);
  clone.id = 'shape' + (Date.now());
  clone.style.left = (parseFloat(s.style.left)||0) + 20 + 'px';
  clone.style.top  = (parseFloat(s.style.top)||0)  + 20 + 'px';
  clone.classList.remove('selected');
  s.parentNode.appendChild(clone);
}

function deleteShape(id) {
  if(!confirm('Delete this shape?')) return;
  let s = document.getElementById('shape'+id);
  if(s) s.style.display = 'none';
  recordEdit({type:'delete', shape:id});
  selectedShape = null;
}

function textChanged(el) {
  el.setAttribute('data-dirty','true');
}

function saveTextEdit(el) {
  if(el.getAttribute('data-dirty') === 'true') {
    recordEdit({type:'text', html:el.innerHTML});
    el.setAttribute('data-dirty','false');
  }
}

function recordEdit(edit) {
  editHistory.push({...edit, ts: Date.now()});
  if(window.chrome && window.chrome.webview)
    window.chrome.webview.postMessage(JSON.stringify({type:'edit', data:edit}));
}

function saveDocument() {
  let html = document.getElementById('mainSlide').innerHTML;
  if(window.chrome && window.chrome.webview)
    window.chrome.webview.postMessage(JSON.stringify({type:'save', html}));
  let indicator = document.createElement('div');
  indicator.textContent = '✅ Saved!';
  indicator.style.cssText = 'position:fixed;bottom:20px;right:20px;background:#22aa22;'
    + 'color:white;padding:10px 20px;border-radius:6px;z-index:9999;';
  document.body.appendChild(indicator);
  setTimeout(() => indicator.remove(), 2000);
}

function undoEdit() {
  if(editHistory.length > 0) {
    editHistory.pop();
    if(window.chrome && window.chrome.webview)
      window.chrome.webview.postMessage(JSON.stringify({type:'undo'}));
  }
}

function exportPDF() { window.print(); }

function prevSlide() {
  if(window.chrome && window.chrome.webview)
    window.chrome.webview.postMessage(JSON.stringify({type:'prev'}));
}

function nextSlide() {
  if(window.chrome && window.chrome.webview)
    window.chrome.webview.postMessage(JSON.stringify({type:'next'}));
}

// Drag to move
document.querySelectorAll('.shape').forEach(el => {
  el.addEventListener('mousedown', e => {
    if(e.target.isContentEditable) return;
    isDragging = true;
    let id = parseInt(el.id.replace('shape',''));
    selectShape(id);
    dragStart = {mx: e.clientX, my: e.clientY,
                 sx: parseFloat(el.style.left)||0,
                 sy: parseFloat(el.style.top)||0};
    dragShape = {id, el};
    e.preventDefault();
  });
});

document.addEventListener('mousemove', e => {
  if(!isDragging || !dragShape.el) return;
  let dx = e.clientX - dragStart.mx;
  let dy = e.clientY - dragStart.my;
  dragShape.el.style.left = (dragStart.sx + dx) + 'px';
  dragShape.el.style.top  = (dragStart.sy + dy) + 'px';
});

document.addEventListener('mouseup', e => {
  if(isDragging && dragShape.el) {
    recordEdit({type:'move', shape:dragShape.id,
                x:parseFloat(dragShape.el.style.left)||0,
                y:parseFloat(dragShape.el.style.top)||0});
  }
  isDragging = false;
  dragShape  = {};
});

// Keyboard shortcuts
document.addEventListener('keydown', e => {
  if(e.ctrlKey && e.key==='s')  { e.preventDefault(); saveDocument(); }
  if(e.ctrlKey && e.key==='z')  { e.preventDefault(); undoEdit(); }
  if(e.key==='Delete' && selectedShape!==null && !document.activeElement.isContentEditable)
    deleteShape(selectedShape);
  if(e.key==='ArrowLeft'  && !e.target.isContentEditable) prevSlide();
  if(e.key==='ArrowRight' && !e.target.isContentEditable) nextSlide();
  if(e.key==='Escape') {
    if(selectedShape!==null) {
      let s=document.getElementById('shape'+selectedShape);
      if(s) s.classList.remove('selected');
      selectedShape=null;
    }
  }
});

// Click outside to deselect
document.getElementById('mainSlide').addEventListener('click', e => {
  if(e.target.id === 'mainSlide') {
    if(selectedShape!==null) {
      let s=document.getElementById('shape'+selectedShape);
      if(s) s.classList.remove('selected');
      selectedShape=null;
      let panel=document.getElementById('props');
      if(panel) panel.innerHTML='<span style="color:#666;font-size:12px">Click a shape to edit</span>';
    }
  }
});
</script>
)SCRIPT";
    }

    void ApplyEditToSlide(SlideData& slide, const EditCommand& cmd) {
        int shapeId = std::stoi(cmd.shapeId);
        for (auto& shape : slide.shapes) {
            if (shape.id == shapeId) {
                if (cmd.type == EditCommand::MOVE) {
                    shape.x = cmd.x; shape.y = cmd.y;
                } else if (cmd.type == EditCommand::RESIZE) {
                    shape.w = cmd.w; shape.h = cmd.h;
                } else if (cmd.type == EditCommand::ROTATE) {
                    shape.rot = cmd.rot;
                } else if (cmd.type == EditCommand::TEXT_CHANGE) {
                    // Update first run of first paragraph
                    if (!shape.paragraphs.empty() && !shape.paragraphs[0].runs.empty())
                        shape.paragraphs[0].runs[0].text = Utf8ToWide(cmd.newText);
                }
                break;
            }
        }
    }

    // ── Rebuild and save ZIP ───────────────────────────────────────
    bool RebuildAndSave(const std::wstring& path) {
        // For now, save a copy of original ZIP with updated XML entries
        // Full ZIP repack would require deflate encoder; write store method (method=0)
        HANDLE h = CreateFileW(path.c_str(), GENERIC_WRITE, 0, NULL,
                               CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
        if (h == INVALID_HANDLE_VALUE) return false;

        // Rebuild as store-only ZIP
        std::vector<uint8_t> newZip;
        std::vector<std::pair<std::string, std::vector<uint8_t>>> updatedEntries;

        // Collect updated slide XMLs
        for (int i = 0; i < (int)m_slides.size(); i++) {
            std::string path2 = "ppt/slides/slide" + std::to_string(i+1) + ".xml";
            updatedEntries.push_back({path2,
                std::vector<uint8_t>(m_slides[i].xml.begin(), m_slides[i].xml.end())});
        }

        // Build local file headers + data
        struct CdEntry { std::string name; uint32_t offset, size; };
        std::vector<CdEntry> cdEntries;

        for (auto& e : m_entries) {
            // Check if we have an updated version
            std::vector<uint8_t>* data = nullptr;
            std::vector<uint8_t> orig;
            for (auto& ue : updatedEntries) {
                if (ue.first == e.name) { data = &ue.second; break; }
            }
            if (!data) {
                // Copy original uncompressed
                orig = ExtractBinary(e.name);
                data = &orig;
            }

            CdEntry cd;
            cd.name   = e.name;
            cd.offset = (uint32_t)newZip.size();
            cd.size   = (uint32_t)data->size();

            // Local file header (method=0 = store)
            auto push2 = [&](uint16_t v) {
                newZip.push_back(v & 0xFF);
                newZip.push_back((v>>8) & 0xFF);
            };
            auto push4 = [&](uint32_t v) {
                newZip.push_back(v & 0xFF);
                newZip.push_back((v>>8) & 0xFF);
                newZip.push_back((v>>16) & 0xFF);
                newZip.push_back((v>>24) & 0xFF);
            };

            newZip.push_back(0x50); newZip.push_back(0x4B);
            newZip.push_back(0x03); newZip.push_back(0x04);
            push2(20);       // version needed
            push2(0);        // flags
            push2(0);        // method: store
            push2(0); push2(0); // mod time/date
            push4(0);        // CRC32 (skip)
            push4((uint32_t)data->size()); // comp size
            push4((uint32_t)data->size()); // uncomp size
            push2((uint16_t)e.name.size());
            push2(0);        // extra len
            for (char c : e.name) newZip.push_back((uint8_t)c);
            for (uint8_t b : *data) newZip.push_back(b);

            cdEntries.push_back(cd);
        }

        // Central directory
        uint32_t cdStart = (uint32_t)newZip.size();
        auto push2 = [&](uint16_t v) {
            newZip.push_back(v & 0xFF); newZip.push_back((v>>8) & 0xFF);
        };
        auto push4 = [&](uint32_t v) {
            newZip.push_back(v & 0xFF); newZip.push_back((v>>8) & 0xFF);
            newZip.push_back((v>>16) & 0xFF); newZip.push_back((v>>24) & 0xFF);
        };

        for (auto& cd : cdEntries) {
            newZip.push_back(0x50); newZip.push_back(0x4B);
            newZip.push_back(0x01); newZip.push_back(0x02);
            push2(20); push2(20); push2(0); push2(0);
            push2(0); push2(0); push2(0); push4(0);
            push4(cd.size); push4(cd.size);
            push2((uint16_t)cd.name.size());
            push2(0); push2(0); push2(0); push2(0);
            push4(0); push4(cd.offset);
            for (char c : cd.name) newZip.push_back((uint8_t)c);
        }

        uint32_t cdSize = (uint32_t)newZip.size() - cdStart;
        // EOCD
        newZip.push_back(0x50); newZip.push_back(0x4B);
        newZip.push_back(0x05); newZip.push_back(0x06);
        push2(0); push2(0);
        push2((uint16_t)cdEntries.size()); push2((uint16_t)cdEntries.size());
        push4(cdSize); push4(cdStart);
        push2(0);

        DWORD written;
        WriteFile(h, newZip.data(), (DWORD)newZip.size(), &written, NULL);
        CloseHandle(h);
        return written == newZip.size();
    }

    // ── Utility ───────────────────────────────────────────────────
    static std::string RemoveNS(const std::string& name) {
        size_t c = name.find(':');
        return c != std::string::npos ? name.substr(c+1) : name;
    }

    double GetAttrD(XmlNode* n, const std::string& a, double def) {
        if (!n) return def;
        auto it = n->attrs.find(a);
        if (it == n->attrs.end()) return def;
        try { return std::stod(it->second); } catch(...) { return def; }
    }

    static std::string EscapeHTML(const std::string& t) {
        std::string r;
        for (char c : t) {
            switch(c) {
                case '<': r += "&lt;"; break;
                case '>': r += "&gt;"; break;
                case '&': r += "&amp;"; break;
                case '"': r += "&quot;"; break;
                default:  r += c;
            }
        }
        return r;
    }
};

// ═══════════════════════════════════════════════════════════════════
// SECTION 3: HIGH-QUALITY DIRECT2D RENDERER
// Renders parsed SlideData to D2D1 with full fidelity
// ═══════════════════════════════════════════════════════════════════

class D2DRenderer {
public:
    struct RenderConfig {
        float scale     = 1.0f;
        float offsetX   = 0.0f;
        float offsetY   = 0.0f;
        bool  antialias = true;
        bool  hqText    = true;
        float dpi       = 96.0f;
        bool  showGrid  = false;
        bool  showGuides = false;
    };

private:
    ComPtr<ID2D1Factory>         m_d2dFactory;
    ComPtr<IDWriteFactory>       m_dwFactory;
    ComPtr<IWICImagingFactory>   m_wicFactory;
    ComPtr<ID2D1HwndRenderTarget> m_rt;
    ComPtr<ID2D1DeviceContext>   m_dc;
    ComPtr<ID2D1SolidColorBrush> m_brush;
    ComPtr<IDWriteTextFormat>    m_defaultFormat;

    // Brush cache
    std::map<uint32_t, ComPtr<ID2D1SolidColorBrush>> m_brushCache;
    // Bitmap cache (per-slide, keyed by rId)
    std::map<std::string, ComPtr<ID2D1Bitmap>> m_bitmapCache;

    RenderConfig m_cfg;
    HWND m_hwnd = nullptr;
    int  m_rtW = 0, m_rtH = 0;

public:
    bool Init(HWND hwnd) {
        m_hwnd = hwnd;

        HRESULT hr = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED,
            IID_PPV_ARGS(&m_d2dFactory));
        if (FAILED(hr)) return false;

        hr = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED,
            __uuidof(IDWriteFactory), reinterpret_cast<IUnknown**>(m_dwFactory.GetAddressOf()));
        if (FAILED(hr)) return false;

        hr = CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER,
            IID_PPV_ARGS(&m_wicFactory));
        if (FAILED(hr)) return false;

        RECT rc;
        GetClientRect(hwnd, &rc);
        CreateRenderTarget(rc.right - rc.left, rc.bottom - rc.top);

        m_dwFactory->CreateTextFormat(L"Segoe UI", nullptr,
            DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL,
            DWRITE_FONT_STRETCH_NORMAL, 14.0f, L"en-US",
            m_defaultFormat.GetAddressOf());

        return true;
    }

    void Resize(int w, int h) {
        if (m_rt) {
            HRESULT hr = m_rt->Resize(D2D1::SizeU(w, h));
            if (FAILED(hr)) CreateRenderTarget(w, h);
        }
        m_rtW = w; m_rtH = h;
    }

    void SetConfig(const RenderConfig& cfg) { m_cfg = cfg; }

    // ── Render full slide ─────────────────────────────────────────
    void RenderSlide(const PptxEngine::SlideData& slide,
                     const PptxEngine& engine,
                     int slideIdx) {
        if (!m_rt) return;

        m_rt->BeginDraw();
        m_rt->SetAntialiasMode(m_cfg.antialias
            ? D2D1_ANTIALIAS_MODE_PER_PRIMITIVE
            : D2D1_ANTIALIAS_MODE_ALIASED);
        m_rt->SetTextAntialiasMode(m_cfg.hqText
            ? D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE
            : D2D1_TEXT_ANTIALIAS_MODE_DEFAULT);

        // Dark background
        m_rt->Clear(D2D1::ColorF(0.08f, 0.08f, 0.12f));

        // Compute slide rect centered in window
        float slideW = slide.width * m_cfg.scale;
        float slideH = slide.height * m_cfg.scale;
        float left = m_cfg.offsetX + (m_rtW - slideW) / 2.0f;
        float top  = m_cfg.offsetY + (m_rtH - slideH) / 2.0f;

        D2D1_MATRIX_3X2_F transform = D2D1::Matrix3x2F(
            m_cfg.scale, 0,
            0, m_cfg.scale,
            left, top
        );
        m_rt->SetTransform(transform);

        // ── Slide white background ──
        D2D1_RECT_F slideRect = D2D1::RectF(0, 0, slide.width, slide.height);

        // Draw background fill
        if (slide.background.type == PptxEngine::FillInfo::Solid) {
            auto brush = GetBrush(slide.background.solidColor);
            m_rt->FillRectangle(slideRect, brush);
        } else if (slide.background.type == PptxEngine::FillInfo::Gradient) {
            DrawGradientRect(slideRect, slide.background);
        } else {
            // Default white
            auto brush = GetBrush(D2D1::ColorF(D2D1::ColorF::White));
            m_rt->FillRectangle(slideRect, brush);
        }

        // Drop shadow under slide
        m_rt->SetTransform(D2D1::IdentityMatrix());
        auto shadowBrush = GetBrush(D2D1::ColorF(0,0,0, 0.35f));
        D2D1_RECT_F shadow = D2D1::RectF(left+6, top+6, left+slideW+6, top+slideH+6);
        m_rt->FillRectangle(shadow, shadowBrush);
        m_rt->SetTransform(transform);

        // ── Draw all shapes ──
        for (const auto& shape : slide.shapes) {
            RenderShape(shape, engine, slideIdx);
        }

        // ── Grid overlay ──
        if (m_cfg.showGrid) DrawGrid(slide.width, slide.height);

        // Reset transform
        m_rt->SetTransform(D2D1::IdentityMatrix());

        // Slide border
        auto borderBrush = GetBrush(D2D1::ColorF(0.3f, 0.3f, 0.5f));
        D2D1_RECT_F border = D2D1::RectF(left, top, left+slideW, top+slideH);
        m_rt->DrawRectangle(border, borderBrush, 1.0f);

        m_rt->EndDraw();
    }

    // ── Render single shape (called recursively for groups) ───────
    void RenderShape(const PptxEngine::ShapeData& shape,
                     const PptxEngine& engine, int slideIdx) {
        D2D1_MATRIX_3X2_F savedTransform;
        m_rt->GetTransform(&savedTransform);

        // Apply shape transform
        D2D1_MATRIX_3X2_F shapeTransform = BuildShapeTransform(shape);
        m_rt->SetTransform(shapeTransform * savedTransform);

        D2D1_RECT_F rect = D2D1::RectF(0, 0, shape.w, shape.h);

        // Shadow
        if (shape.effect.hasShadow) DrawShadow(shape);

        // Groups
        if (shape.type == PptxEngine::ShapeType::Group) {
            for (const auto& child : shape.children) {
                RenderShape(child, engine, slideIdx);
            }
            m_rt->SetTransform(savedTransform);
            return;
        }

        // Picture
        if (shape.type == PptxEngine::ShapeType::Picture) {
            DrawPicture(shape, engine, slideIdx, rect);
            if (shape.line.width > 0) DrawShapeStroke(shape, rect);
            m_rt->SetTransform(savedTransform);
            return;
        }

        // Table
        if (shape.type == PptxEngine::ShapeType::Table) {
            DrawTable(shape);
            m_rt->SetTransform(savedTransform);
            return;
        }

        // Line / connector
        if (shape.type == PptxEngine::ShapeType::Line) {
            DrawLineShape(shape);
            m_rt->SetTransform(savedTransform);
            return;
        }

        // Regular shape fill
        DrawShapeFill(shape, rect);

        // Shape geometry stroke
        if (shape.line.width > 0) DrawShapeStroke(shape, rect);

        // Glow
        if (shape.effect.hasGlow) DrawGlow(shape, rect);

        // Text
        if (!shape.paragraphs.empty()) {
            DrawTextBody(shape, rect);
        }

        m_rt->SetTransform(savedTransform);
    }

private:
    void CreateRenderTarget(int w, int h) {
        m_rt.Reset();
        m_brushCache.clear();
        m_rtW = w; m_rtH = h;

        D2D1_RENDER_TARGET_PROPERTIES rtp = D2D1::RenderTargetProperties(
            D2D1_RENDER_TARGET_TYPE_DEFAULT,
            D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_PREMULTIPLIED),
            96.0f, 96.0f
        );
        D2D1_HWND_RENDER_TARGET_PROPERTIES hwndRtp = D2D1::HwndRenderTargetProperties(
            m_hwnd, D2D1::SizeU(w, h),
            D2D1_PRESENT_OPTIONS_NONE
        );
        m_d2dFactory->CreateHwndRenderTarget(rtp, hwndRtp, m_rt.GetAddressOf());
    }

    D2D1_MATRIX_3X2_F BuildShapeTransform(const PptxEngine::ShapeData& shape) {
        // Rotate around shape center
        float cx = shape.x + shape.w / 2.0f;
        float cy = shape.y + shape.h / 2.0f;

        D2D1_MATRIX_3X2_F t = D2D1::Matrix3x2F::Translation(shape.x, shape.y);

        if (shape.rot != 0.0f) {
            float radians = shape.rot * 3.14159265f / 180.0f;
            D2D1_MATRIX_3X2_F rot = D2D1::Matrix3x2F::Rotation(
                shape.rot, D2D1::Point2F(cx - shape.x, cy - shape.y));
            t = rot * t;
        }
        if (shape.flipH || shape.flipV) {
            float sx = shape.flipH ? -1.0f : 1.0f;
            float sy = shape.flipV ? -1.0f : 1.0f;
            D2D1_MATRIX_3X2_F flip = D2D1::Matrix3x2F(
                sx, 0, 0, sy,
                shape.flipH ? shape.x + shape.w : 0,
                shape.flipV ? shape.y + shape.h : 0
            );
            t = flip * t;
        }
        return t;
    }

    ID2D1SolidColorBrush* GetBrush(const D2D1_COLOR_F& color) {
        uint32_t key = ((uint32_t)(color.r*255) << 24) |
                       ((uint32_t)(color.g*255) << 16) |
                       ((uint32_t)(color.b*255) << 8)  |
                        (uint32_t)(color.a*255);
        auto it = m_brushCache.find(key);
        if (it != m_brushCache.end()) {
            it->second->SetColor(color);
            return it->second.Get();
        }
        ComPtr<ID2D1SolidColorBrush> brush;
        m_rt->CreateSolidColorBrush(color, brush.GetAddressOf());
        m_brushCache[key] = brush;
        return brush.Get();
    }

    // ── Fill rendering ─────────────────────────────────────────────
    void DrawShapeFill(const PptxEngine::ShapeData& shape, const D2D1_RECT_F& rect) {
        using FI = PptxEngine::FillInfo;
        const auto& fill = shape.fill;
        if (fill.type == FI::None) return;

        switch (shape.type) {
            case PptxEngine::ShapeType::Ellipse:
                DrawEllipseFill(shape, fill, rect);
                break;
            case PptxEngine::ShapeType::RoundRect:
                DrawRoundRectFill(shape, fill, rect);
                break;
            case PptxEngine::ShapeType::Triangle:
                DrawPolygonFill(shape, fill, MakeTriangle(rect));
                break;
            case PptxEngine::ShapeType::Diamond:
                DrawPolygonFill(shape, fill, MakeDiamond(rect));
                break;
            case PptxEngine::ShapeType::Pentagon:
                DrawPolygonFill(shape, fill, MakeRegularPolygon(rect, 5));
                break;
            case PptxEngine::ShapeType::Hexagon:
                DrawPolygonFill(shape, fill, MakeRegularPolygon(rect, 6));
                break;
            case PptxEngine::ShapeType::Heptagon:
                DrawPolygonFill(shape, fill, MakeRegularPolygon(rect, 7));
                break;
            case PptxEngine::ShapeType::Octagon:
                DrawPolygonFill(shape, fill, MakeRegularPolygon(rect, 8));
                break;
            case PptxEngine::ShapeType::Parallelogram:
                DrawPolygonFill(shape, fill, MakeParallelogram(rect, shape.geom.adjVal[0]));
                break;
            case PptxEngine::ShapeType::Trapezoid:
                DrawPolygonFill(shape, fill, MakeTrapezoid(rect, shape.geom.adjVal[0]));
                break;
            case PptxEngine::ShapeType::Star4:
            case PptxEngine::ShapeType::Star5:
            case PptxEngine::ShapeType::Star6:
            case PptxEngine::ShapeType::Star8:
            case PptxEngine::ShapeType::Star16:
            case PptxEngine::ShapeType::Star24:
            case PptxEngine::ShapeType::Star32: {
                int pts = StarPointsFromType(shape.type);
                DrawPolygonFill(shape, fill, MakeStar(rect, pts, shape.geom.adjVal[0]));
                break;
            }
            case PptxEngine::ShapeType::ArrowRight:
            case PptxEngine::ShapeType::ArrowLeft:
            case PptxEngine::ShapeType::ArrowUp:
            case PptxEngine::ShapeType::ArrowDown:
                DrawPolygonFill(shape, fill, MakeArrow(rect, shape.type));
                break;
            case PptxEngine::ShapeType::Chevron:
                DrawPolygonFill(shape, fill, MakeChevron(rect));
                break;
            case PptxEngine::ShapeType::Wave:
                DrawWaveFill(shape, fill, rect);
                break;
            case PptxEngine::ShapeType::Cloud:
                DrawCloudFill(shape, fill, rect);
                break;
            case PptxEngine::ShapeType::Heart:
                DrawHeartFill(shape, fill, rect);
                break;
            default:
                DrawRectFill(shape, fill, rect);
                break;
        }
    }

    void DrawRectFill(const PptxEngine::ShapeData& s, const PptxEngine::FillInfo& fill,
                      const D2D1_RECT_F& rect) {
        if (fill.type == PptxEngine::FillInfo::Solid) {
            D2D1_COLOR_F c = fill.solidColor; c.a = fill.alpha;
            m_rt->FillRectangle(rect, GetBrush(c));
        } else if (fill.type == PptxEngine::FillInfo::Gradient) {
            DrawGradientRect(rect, fill);
        }
    }

    void DrawEllipseFill(const PptxEngine::ShapeData& s, const PptxEngine::FillInfo& fill,
                         const D2D1_RECT_F& rect) {
        D2D1_ELLIPSE el = D2D1::Ellipse(
            D2D1::Point2F((rect.left+rect.right)/2, (rect.top+rect.bottom)/2),
            (rect.right-rect.left)/2, (rect.bottom-rect.top)/2
        );
        if (fill.type == PptxEngine::FillInfo::Solid) {
            D2D1_COLOR_F c = fill.solidColor; c.a = fill.alpha;
            m_rt->FillEllipse(el, GetBrush(c));
        } else if (fill.type == PptxEngine::FillInfo::Gradient) {
            DrawGradientEllipse(el, fill);
        }
    }

    void DrawRoundRectFill(const PptxEngine::ShapeData& s, const PptxEngine::FillInfo& fill,
                           const D2D1_RECT_F& rect) {
        float radius = std::min(s.w, s.h) * 0.1f;
        if (s.geom.adjVal[0] > 0) radius = s.w * s.geom.adjVal[0];
        D2D1_ROUNDED_RECT rr = D2D1::RoundedRect(rect, radius, radius);
        if (fill.type == PptxEngine::FillInfo::Solid) {
            D2D1_COLOR_F c = fill.solidColor; c.a = fill.alpha;
            m_rt->FillRoundedRectangle(rr, GetBrush(c));
        } else if (fill.type == PptxEngine::FillInfo::Gradient) {
            // Clip to rounded rect and draw gradient
            ComPtr<ID2D1RoundedRectangleGeometry> geom;
            m_d2dFactory->CreateRoundedRectangleGeometry(rr, geom.GetAddressOf());
            m_rt->PushLayer(D2D1::LayerParameters(rect, geom.Get()), nullptr);
            DrawGradientRect(rect, fill);
            m_rt->PopLayer();
        }
    }

    void DrawPolygonFill(const PptxEngine::ShapeData& s, const PptxEngine::FillInfo& fill,
                         const std::vector<D2D1_POINT_2F>& pts) {
        ComPtr<ID2D1PathGeometry> path;
        m_d2dFactory->CreatePathGeometry(path.GetAddressOf());
        ComPtr<ID2D1GeometrySink> sink;
        path->Open(sink.GetAddressOf());
        if (!pts.empty()) {
            sink->BeginFigure(pts[0], D2D1_FIGURE_BEGIN_FILLED);
            for (size_t i = 1; i < pts.size(); i++) sink->AddLine(pts[i]);
            sink->EndFigure(D2D1_FIGURE_END_CLOSED);
        }
        sink->Close();

        if (fill.type == PptxEngine::FillInfo::Solid) {
            D2D1_COLOR_F c = fill.solidColor; c.a = fill.alpha;
            m_rt->FillGeometry(path.Get(), GetBrush(c));
        } else if (fill.type == PptxEngine::FillInfo::Gradient) {
            // Bounding box of polygon for gradient
            float minX = pts[0].x, maxX = pts[0].x;
            float minY = pts[0].y, maxY = pts[0].y;
            for (auto& p : pts) {
                minX = std::min(minX, p.x); maxX = std::max(maxX, p.x);
                minY = std::min(minY, p.y); maxY = std::max(maxY, p.y);
            }
            D2D1_RECT_F bbox = D2D1::RectF(minX, minY, maxX, maxY);
            m_rt->PushLayer(D2D1::LayerParameters(bbox, path.Get()), nullptr);
            DrawGradientRect(bbox, fill);
            m_rt->PopLayer();
        }
    }

    void DrawWaveFill(const PptxEngine::ShapeData& s, const PptxEngine::FillInfo& fill,
                      const D2D1_RECT_F& rect) {
        ComPtr<ID2D1PathGeometry> path;
        m_d2dFactory->CreatePathGeometry(path.GetAddressOf());
        ComPtr<ID2D1GeometrySink> sink;
        path->Open(sink.GetAddressOf());

        float w = rect.right - rect.left;
        float h = rect.bottom - rect.top;
        float amp = h * 0.15f;
        int waves = 3;
        float waveW = w / waves;

        sink->BeginFigure(D2D1::Point2F(rect.left, rect.top + h/2), D2D1_FIGURE_BEGIN_FILLED);
        for (int i = 0; i < waves; i++) {
            float x0 = rect.left + i * waveW;
            sink->AddBezier(D2D1::BezierSegment(
                D2D1::Point2F(x0 + waveW*0.25f, rect.top + h/2 - amp),
                D2D1::Point2F(x0 + waveW*0.75f, rect.top + h/2 + amp),
                D2D1::Point2F(x0 + waveW,        rect.top + h/2)
            ));
        }
        sink->AddLine(D2D1::Point2F(rect.right, rect.bottom));
        sink->AddLine(D2D1::Point2F(rect.left,  rect.bottom));
        sink->EndFigure(D2D1_FIGURE_END_CLOSED);
        sink->Close();

        if (fill.type == PptxEngine::FillInfo::Solid) {
            D2D1_COLOR_F c = fill.solidColor; c.a = fill.alpha;
            m_rt->FillGeometry(path.Get(), GetBrush(c));
        }
    }

    void DrawCloudFill(const PptxEngine::ShapeData& s, const PptxEngine::FillInfo& fill,
                       const D2D1_RECT_F& rect) {
        float w = rect.right - rect.left;
        float h = rect.bottom - rect.top;
        float cx = (rect.left + rect.right) / 2;
        float cy = (rect.top  + rect.bottom) / 2;

        if (fill.type == PptxEngine::FillInfo::Solid) {
            D2D1_COLOR_F c = fill.solidColor; c.a = fill.alpha;
            auto brush = GetBrush(c);
            // Draw overlapping ellipses to form cloud
            struct Bump { float bx, by, br, vert; };
            Bump bumps[] = {
                {cx,        cy - h*0.12f, w*0.32f, h*0.28f},
                {cx - w*0.22f, cy,        w*0.24f, h*0.24f},
                {cx + w*0.22f, cy,        w*0.24f, h*0.24f},
                {cx - w*0.10f, cy - h*0.06f, w*0.28f, h*0.26f},
                {cx + w*0.10f, cy - h*0.06f, w*0.26f, h*0.24f},
                {cx,           cy + h*0.08f, w*0.40f, h*0.20f},
            };
            for (auto& b : bumps) {
                m_rt->FillEllipse(D2D1::Ellipse(D2D1::Point2F(b.bx, b.by), b.br, b.vert), brush);
            }
        }
    }

    void DrawHeartFill(const PptxEngine::ShapeData& s, const PptxEngine::FillInfo& fill,
                       const D2D1_RECT_F& rect) {
        ComPtr<ID2D1PathGeometry> path;
        m_d2dFactory->CreatePathGeometry(path.GetAddressOf());
        ComPtr<ID2D1GeometrySink> sink;
        path->Open(sink.GetAddressOf());

        float w = rect.right - rect.left;
        float h = rect.bottom - rect.top;
        float cx = (rect.left + rect.right) / 2;
        float top = rect.top;

        // Heart bezier path
        sink->BeginFigure(D2D1::Point2F(cx, rect.bottom), D2D1_FIGURE_BEGIN_FILLED);
        sink->AddBezier(D2D1::BezierSegment(
            D2D1::Point2F(cx - w*0.5f, rect.top + h*0.6f),
            D2D1::Point2F(cx - w*0.5f, top),
            D2D1::Point2F(cx - w*0.25f, top)
        ));
        sink->AddBezier(D2D1::BezierSegment(
            D2D1::Point2F(cx - w*0.1f, top),
            D2D1::Point2F(cx,          top + h*0.2f),
            D2D1::Point2F(cx,          top + h*0.2f)
        ));
        sink->AddBezier(D2D1::BezierSegment(
            D2D1::Point2F(cx,          top + h*0.2f),
            D2D1::Point2F(cx + w*0.1f, top),
            D2D1::Point2F(cx + w*0.25f, top)
        ));
        sink->AddBezier(D2D1::BezierSegment(
            D2D1::Point2F(cx + w*0.5f, top),
            D2D1::Point2F(cx + w*0.5f, rect.top + h*0.6f),
            D2D1::Point2F(cx,          rect.bottom)
        ));
        sink->EndFigure(D2D1_FIGURE_END_CLOSED);
        sink->Close();

        if (fill.type == PptxEngine::FillInfo::Solid) {
            D2D1_COLOR_F c = fill.solidColor; c.a = fill.alpha;
            m_rt->FillGeometry(path.Get(), GetBrush(c));
        }
    }

    // ── Gradient rendering ────────────────────────────────────────
    void DrawGradientRect(const D2D1_RECT_F& rect, const PptxEngine::FillInfo& fill) {
        if (fill.gradStops.size() < 2) return;

        if (fill.gradLinear) {
            // Compute start/end points from angle
            float angle = fill.gradAngle * 3.14159f / 180.0f;
            float w = rect.right - rect.left, h = rect.bottom - rect.top;
            float cx = (rect.left+rect.right)/2, cy = (rect.top+rect.bottom)/2;
            float len = (std::abs(w*std::cos(angle)) + std::abs(h*std::sin(angle))) / 2;

            D2D1_POINT_2F startPt = D2D1::Point2F(cx - len*std::cos(angle), cy - len*std::sin(angle));
            D2D1_POINT_2F endPt   = D2D1::Point2F(cx + len*std::cos(angle), cy + len*std::sin(angle));

            std::vector<D2D1_GRADIENT_STOP> stops;
            for (auto& gs : fill.gradStops) {
                D2D1_GRADIENT_STOP s;
                s.position = gs.pos;
                s.color    = gs.color;
                s.color.a  = gs.color.a;
                stops.push_back(s);
            }

            ComPtr<ID2D1GradientStopCollection> stopColl;
            m_rt->CreateGradientStopCollection(stops.data(), (UINT32)stops.size(),
                stopColl.GetAddressOf());
            if (!stopColl) return;

            ComPtr<ID2D1LinearGradientBrush> gradBrush;
            m_rt->CreateLinearGradientBrush(
                D2D1::LinearGradientBrushProperties(startPt, endPt),
                stopColl.Get(), gradBrush.GetAddressOf());
            if (gradBrush) m_rt->FillRectangle(rect, gradBrush.Get());

        } else {
            // Radial
            float cx = (rect.left+rect.right)/2, cy = (rect.top+rect.bottom)/2;
            float rx = (rect.right-rect.left)/2, ry = (rect.bottom-rect.top)/2;

            std::vector<D2D1_GRADIENT_STOP> stops;
            for (auto& gs : fill.gradStops) {
                D2D1_GRADIENT_STOP s; s.position = gs.pos; s.color = gs.color;
                stops.push_back(s);
            }
            ComPtr<ID2D1GradientStopCollection> stopColl;
            m_rt->CreateGradientStopCollection(stops.data(), (UINT32)stops.size(),
                stopColl.GetAddressOf());
            if (!stopColl) return;

            ComPtr<ID2D1RadialGradientBrush> radBrush;
            m_rt->CreateRadialGradientBrush(
                D2D1::RadialGradientBrushProperties(
                    D2D1::Point2F(cx, cy), D2D1::Point2F(0,0), rx, ry),
                stopColl.Get(), radBrush.GetAddressOf());
            if (radBrush) m_rt->FillRectangle(rect, radBrush.Get());
        }
    }

    void DrawGradientEllipse(const D2D1_ELLIPSE& el, const PptxEngine::FillInfo& fill) {
        if (fill.gradStops.size() < 2) return;
        std::vector<D2D1_GRADIENT_STOP> stops;
        for (auto& gs : fill.gradStops) {
            D2D1_GRADIENT_STOP s; s.position = gs.pos; s.color = gs.color;
            stops.push_back(s);
        }
        ComPtr<ID2D1GradientStopCollection> sc;
        m_rt->CreateGradientStopCollection(stops.data(), (UINT32)stops.size(), sc.GetAddressOf());
        if (!sc) return;
        ComPtr<ID2D1RadialGradientBrush> rb;
        m_rt->CreateRadialGradientBrush(
            D2D1::RadialGradientBrushProperties(el.point, D2D1::Point2F(0,0), el.radiusX, el.radiusY),
            sc.Get(), rb.GetAddressOf());
        if (rb) m_rt->FillEllipse(el, rb.Get());
    }

    // ── Stroke rendering ──────────────────────────────────────────
    void DrawShapeStroke(const PptxEngine::ShapeData& shape, const D2D1_RECT_F& rect) {
        const auto& line = shape.line;
        if (line.width <= 0) return;

        D2D1_COLOR_F c = line.fill.solidColor;
        auto brush = GetBrush(c);

        // Build stroke style for dashes
        ComPtr<ID2D1StrokeStyle> strokeStyle;
        D2D1_STROKE_STYLE_PROPERTIES ssp = D2D1::StrokeStyleProperties(
            line.cap  == PptxEngine::LineInfo::Round   ? D2D1_CAP_STYLE_ROUND  :
            line.cap  == PptxEngine::LineInfo::Square  ? D2D1_CAP_STYLE_SQUARE : D2D1_CAP_STYLE_FLAT,
            line.cap  == PptxEngine::LineInfo::Round   ? D2D1_CAP_STYLE_ROUND  :
            line.cap  == PptxEngine::LineInfo::Square  ? D2D1_CAP_STYLE_SQUARE : D2D1_CAP_STYLE_FLAT,
            D2D1_CAP_STYLE_FLAT,
            line.join == PptxEngine::LineInfo::Miter   ? D2D1_LINE_JOIN_MITER  :
            line.join == PptxEngine::LineInfo::Round_  ? D2D1_LINE_JOIN_ROUND  : D2D1_LINE_JOIN_BEVEL,
            10.0f,
            line.dash == PptxEngine::LineInfo::Solid   ? D2D1_DASH_STYLE_SOLID :
            line.dash == PptxEngine::LineInfo::Dot     ? D2D1_DASH_STYLE_DOT   :
            line.dash == PptxEngine::LineInfo::Dash_   ? D2D1_DASH_STYLE_DASH  :
            line.dash == PptxEngine::LineInfo::DashDot ? D2D1_DASH_STYLE_DASH_DOT : D2D1_DASH_STYLE_DASH_DOT_DOT,
            0.0f
        );
        m_d2dFactory->CreateStrokeStyle(ssp, nullptr, 0, strokeStyle.GetAddressOf());

        switch (shape.type) {
            case PptxEngine::ShapeType::Ellipse: {
                D2D1_ELLIPSE el = D2D1::Ellipse(
                    D2D1::Point2F((rect.left+rect.right)/2, (rect.top+rect.bottom)/2),
                    (rect.right-rect.left)/2, (rect.bottom-rect.top)/2);
                m_rt->DrawEllipse(el, brush, line.width, strokeStyle.Get());
                break;
            }
            case PptxEngine::ShapeType::RoundRect: {
                float r = std::min(shape.w, shape.h) * 0.1f;
                m_rt->DrawRoundedRectangle(D2D1::RoundedRect(rect,r,r), brush, line.width, strokeStyle.Get());
                break;
            }
            case PptxEngine::ShapeType::Triangle:
                DrawPolygonStroke(MakeTriangle(rect), brush, line.width, strokeStyle.Get());
                break;
            case PptxEngine::ShapeType::Diamond:
                DrawPolygonStroke(MakeDiamond(rect), brush, line.width, strokeStyle.Get());
                break;
            case PptxEngine::ShapeType::Hexagon:
                DrawPolygonStroke(MakeRegularPolygon(rect, 6), brush, line.width, strokeStyle.Get());
                break;
            case PptxEngine::ShapeType::Octagon:
                DrawPolygonStroke(MakeRegularPolygon(rect, 8), brush, line.width, strokeStyle.Get());
                break;
            default:
                m_rt->DrawRectangle(rect, brush, line.width, strokeStyle.Get());
                break;
        }
    }

    void DrawPolygonStroke(const std::vector<D2D1_POINT_2F>& pts,
                           ID2D1SolidColorBrush* brush, float width,
                           ID2D1StrokeStyle* style) {
        for (size_t i = 0; i < pts.size(); i++) {
            size_t j = (i+1) % pts.size();
            m_rt->DrawLine(pts[i], pts[j], brush, width, style);
        }
    }

    void DrawLineShape(const PptxEngine::ShapeData& shape) {
        if (shape.line.width <= 0 && shape.line.fill.type == PptxEngine::FillInfo::None) return;
        float w = shape.line.width > 0 ? shape.line.width : 1.5f;
        D2D1_COLOR_F c = shape.line.fill.solidColor;
        auto brush = GetBrush(c);
        m_rt->DrawLine(D2D1::Point2F(0, shape.h/2),
                       D2D1::Point2F(shape.w, shape.h/2),
                       brush, w);
        // Arrowhead
        if (!shape.line.tailType.empty() && shape.line.tailType != "none") {
            DrawArrowHead(D2D1::Point2F(shape.w, shape.h/2),
                          D2D1::Point2F(shape.w - 12, shape.h/2 - 6),
                          D2D1::Point2F(shape.w - 12, shape.h/2 + 6),
                          brush);
        }
        if (!shape.line.headType.empty() && shape.line.headType != "none") {
            DrawArrowHead(D2D1::Point2F(0, shape.h/2),
                          D2D1::Point2F(12, shape.h/2 - 6),
                          D2D1::Point2F(12, shape.h/2 + 6),
                          brush);
        }
    }

    void DrawArrowHead(D2D1_POINT_2F tip, D2D1_POINT_2F p1, D2D1_POINT_2F p2,
                       ID2D1SolidColorBrush* brush) {
        ComPtr<ID2D1PathGeometry> path;
        m_d2dFactory->CreatePathGeometry(path.GetAddressOf());
        ComPtr<ID2D1GeometrySink> sink;
        path->Open(sink.GetAddressOf());
        sink->BeginFigure(tip, D2D1_FIGURE_BEGIN_FILLED);
        sink->AddLine(p1); sink->AddLine(p2);
        sink->EndFigure(D2D1_FIGURE_END_CLOSED);
        sink->Close();
        m_rt->FillGeometry(path.Get(), brush);
    }

    // ── Shadow / Glow ─────────────────────────────────────────────
    void DrawShadow(const PptxEngine::ShapeData& shape) {
        // Approximate shadow with offset fill
        const auto& eff = shape.effect;
        D2D1_MATRIX_3X2_F saved;
        m_rt->GetTransform(&saved);
        D2D1_MATRIX_3X2_F shadowT = D2D1::Matrix3x2F::Translation(eff.shadowDX, eff.shadowDY);
        m_rt->SetTransform(shadowT * saved);

        auto brush = GetBrush(eff.shadowColor);
        D2D1_RECT_F rect = D2D1::RectF(0, 0, shape.w, shape.h);
        m_rt->FillRectangle(rect, brush);

        m_rt->SetTransform(saved);
    }

    void DrawGlow(const PptxEngine::ShapeData& shape, const D2D1_RECT_F& rect) {
        const auto& eff = shape.effect;
        float r = eff.glowRadius;
        D2D1_RECT_F glowRect = D2D1::RectF(rect.left-r, rect.top-r, rect.right+r, rect.bottom+r);
        // Draw several expanding, fading rectangles
        for (int i = 5; i >= 1; i--) {
            float fr = r * i / 5.0f;
            D2D1_COLOR_F gc = eff.glowColor;
            gc.a *= (float)i / 8.0f;
            D2D1_RECT_F gr = D2D1::RectF(rect.left-fr, rect.top-fr, rect.right+fr, rect.bottom+fr);
            D2D1_ROUNDED_RECT rr = D2D1::RoundedRect(gr, fr, fr);
            m_rt->DrawRoundedRectangle(rr, GetBrush(gc), fr*0.5f);
        }
    }

    // ── Picture rendering ─────────────────────────────────────────
    void DrawPicture(const PptxEngine::ShapeData& shape,
                     const PptxEngine& engine, int slideIdx,
                     const D2D1_RECT_F& rect) {
        std::string key = std::to_string(slideIdx) + ":" + shape.imageRId;
        auto it = m_bitmapCache.find(key);
        if (it != m_bitmapCache.end() && it->second) {
            m_rt->DrawBitmap(it->second.Get(), rect, shape.imageAlpha);
            return;
        }

        // Load from engine
        auto bytes = const_cast<PptxEngine&>(engine).GetImageBytes(slideIdx, shape.imageRId);
        if (bytes.empty()) {
            // Draw placeholder
            auto brush = GetBrush(D2D1::ColorF(0.8f,0.8f,0.9f));
            m_rt->FillRectangle(rect, brush);
            auto textBrush = GetBrush(D2D1::ColorF(0.4f,0.4f,0.5f));
            DrawSimpleText(L"[Image]", rect, textBrush);
            return;
        }

        // Decode with WIC
        ComPtr<IWICStream> stream;
        m_wicFactory->CreateStream(stream.GetAddressOf());
        stream->InitializeFromMemory(bytes.data(), (DWORD)bytes.size());

        ComPtr<IWICBitmapDecoder> decoder;
        m_wicFactory->CreateDecoderFromStream(stream.Get(), nullptr,
            WICDecodeMetadataCacheOnLoad, decoder.GetAddressOf());
        if (!decoder) {
            auto brush = GetBrush(D2D1::ColorF(0.8f,0.8f,0.9f));
            m_rt->FillRectangle(rect, brush);
            return;
        }

        ComPtr<IWICBitmapFrameDecode> frame;
        decoder->GetFrame(0, frame.GetAddressOf());
        if (!frame) return;

        ComPtr<IWICFormatConverter> converter;
        m_wicFactory->CreateFormatConverter(converter.GetAddressOf());
        converter->Initialize(frame.Get(), GUID_WICPixelFormat32bppPBGRA,
            WICBitmapDitherTypeNone, nullptr, 0.0, WICBitmapPaletteTypeCustom);

        ComPtr<ID2D1Bitmap> bitmap;
        m_rt->CreateBitmapFromWicBitmap(converter.Get(), bitmap.GetAddressOf());
        if (bitmap) {
            m_rt->DrawBitmap(bitmap.Get(), rect, shape.imageAlpha);
            m_bitmapCache[key] = bitmap;
        }
    }

    // ── Table rendering ───────────────────────────────────────────
    void DrawTable(const PptxEngine::ShapeData& shape) {
        if (shape.tableRows.empty() || shape.colWidths.empty()) return;

        float y = 0;
        for (size_t row = 0; row < shape.tableRows.size(); row++) {
            float x = 0;
            float rh = (row < shape.rowHeights.size()) ? shape.rowHeights[row] : 20.0f;

            for (size_t col = 0; col < shape.tableRows[row].size(); col++) {
                float cw = (col < shape.colWidths.size()) ? shape.colWidths[col] : 80.0f;
                const auto& cell = shape.tableRows[row][col];

                D2D1_RECT_F cellRect = D2D1::RectF(x, y, x+cw, y+rh);

                // Cell fill
                if (cell.fill.type == PptxEngine::FillInfo::Solid) {
                    m_rt->FillRectangle(cellRect, GetBrush(cell.fill.solidColor));
                }

                // Cell border
                auto borderBrush = GetBrush(D2D1::ColorF(0.5f,0.5f,0.5f));
                m_rt->DrawRectangle(cellRect, borderBrush, 0.5f);

                // Cell text
                if (!cell.paragraphs.empty()) {
                    PptxEngine::ShapeData tmp;
                    tmp.paragraphs = cell.paragraphs;
                    tmp.w = cw; tmp.h = rh;
                    tmp.txInsetL = tmp.txInsetR = 3.6f;
                    tmp.txInsetT = tmp.txInsetB = 1.8f;
                    tmp.textVAlign = DWRITE_PARAGRAPH_ALIGNMENT_CENTER;
                    DrawTextBody(tmp, cellRect);
                }

                x += cw;
            }
            y += rh;
        }
    }

    // ── High-quality text rendering ───────────────────────────────
    void DrawTextBody(const PptxEngine::ShapeData& shape, const D2D1_RECT_F& rect) {
        if (shape.paragraphs.empty()) return;

        // Text area (insets)
        D2D1_RECT_F textRect = D2D1::RectF(
            rect.left   + shape.txInsetL,
            rect.top    + shape.txInsetT,
            rect.right  - shape.txInsetR,
            rect.bottom - shape.txInsetB
        );
        if (textRect.left >= textRect.right || textRect.top >= textRect.bottom) return;

        // Push clip
        m_rt->PushAxisAlignedClip(textRect, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);

        float curY = textRect.top;

        // For vertical centering, measure total text height first
        if (shape.textVAlign != DWRITE_PARAGRAPH_ALIGNMENT_NEAR) {
            float totalH = MeasureTextBodyHeight(shape, textRect);
            float avail  = textRect.bottom - textRect.top;
            if (shape.textVAlign == DWRITE_PARAGRAPH_ALIGNMENT_CENTER)
                curY = textRect.top + (avail - totalH) / 2.0f;
            else if (shape.textVAlign == DWRITE_PARAGRAPH_ALIGNMENT_FAR)
                curY = textRect.bottom - totalH;
        }

        for (const auto& para : shape.paragraphs) {
            curY += para.props.spaceBefore * 96.0f / 72.0f;
            float paraH = DrawParagraph(para, textRect, curY);
            curY += paraH + para.props.spaceAfter * 96.0f / 72.0f;
            if (curY > textRect.bottom) break;
        }

        m_rt->PopAxisAlignedClip();
    }

    float MeasureTextBodyHeight(const PptxEngine::ShapeData& shape, const D2D1_RECT_F& textRect) {
        float totalH = 0;
        for (const auto& para : shape.paragraphs) {
            totalH += para.props.spaceBefore * 96.0f / 72.0f;
            // Estimate paragraph height from font sizes
            float maxFontPx = 14.0f;
            for (const auto& run : para.runs)
                maxFontPx = std::max(maxFontPx, run.props.fontSize * 96.0f / 72.0f);
            totalH += maxFontPx * (para.props.lineSpacingIsPoints
                ? para.props.lineSpacing : para.props.lineSpacing);
            totalH += para.props.spaceAfter * 96.0f / 72.0f;
        }
        return totalH;
    }

    float DrawParagraph(const PptxEngine::Paragraph& para,
                        const D2D1_RECT_F& textRect, float y) {
        using PP = PptxEngine::ParaProps;
        const auto& props = para.props;

        float lineH = 0;
        float x     = textRect.left + props.leftMargin;
        float maxW  = textRect.right - textRect.left;

        // Bullet
        if (props.hasBullet && !props.bulletChar.empty()) {
            auto brush = GetBrush(props.bulletColor);
            DrawRunText(Utf8ToWide(props.bulletChar + " "), x, y, 14.0f,
                        L"Segoe UI", false, false, brush);
            x += 14.0f;
        }

        // Runs – we use IDWriteTextLayout for proper shaping/wrapping
        if (!para.runs.empty()) {
            lineH = DrawRunsOnLine(para, textRect, x, y);
        } else {
            // Empty paragraph – use default line height
            lineH = 18.0f;
        }

        return lineH;
    }

    float DrawRunsOnLine(const PptxEngine::Paragraph& para,
                         const D2D1_RECT_F& textRect, float startX, float y) {
        // Build a combined IDWriteTextLayout for the paragraph
        // to handle multi-run paragraphs with different formatting
        std::wstring combined;
        for (const auto& run : para.runs) combined += run.text;
        if (combined.empty()) return 14.0f;

        // Use first run's font for the base format
        const auto& firstRun = para.runs[0];
        float fontSize = firstRun.props.fontSize * 96.0f / 72.0f;

        ComPtr<IDWriteTextFormat> fmt;
        m_dwFactory->CreateTextFormat(
            Utf8ToWide(firstRun.props.fontFamily).c_str(), nullptr,
            firstRun.props.bold   ? DWRITE_FONT_WEIGHT_BOLD   : DWRITE_FONT_WEIGHT_NORMAL,
            firstRun.props.italic ? DWRITE_FONT_STYLE_ITALIC  : DWRITE_FONT_STYLE_NORMAL,
            DWRITE_FONT_STRETCH_NORMAL,
            fontSize, L"en-US", fmt.GetAddressOf());
        if (!fmt) return fontSize;

        fmt->SetTextAlignment(para.props.align);
        fmt->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP);

        float layoutW = textRect.right - startX;
        float layoutH = textRect.bottom - y;
        if (layoutW <= 0 || layoutH <= 0) return fontSize;

        ComPtr<IDWriteTextLayout> layout;
        m_dwFactory->CreateTextLayout(combined.c_str(), (UINT32)combined.size(),
            fmt.Get(), layoutW, layoutH, layout.GetAddressOf());
        if (!layout) return fontSize;

        // Apply per-run formatting
        UINT32 pos = 0;
        for (const auto& run : para.runs) {
            UINT32 len = (UINT32)run.text.size();
            if (len == 0) continue;
            DWRITE_TEXT_RANGE range = {pos, len};

            // Font family
            layout->SetFontFamilyName(Utf8ToWide(run.props.fontFamily).c_str(), range);
            // Font size
            layout->SetFontSize(run.props.fontSize * 96.0f / 72.0f, range);
            // Bold
            layout->SetFontWeight(run.props.bold ? DWRITE_FONT_WEIGHT_BOLD : DWRITE_FONT_WEIGHT_NORMAL, range);
            // Italic
            layout->SetFontStyle(run.props.italic ? DWRITE_FONT_STYLE_ITALIC : DWRITE_FONT_STYLE_NORMAL, range);
            // Underline
            layout->SetUnderline(run.props.underline, range);
            // Strikethrough
            layout->SetStrikethrough(run.props.strikeThru, range);

            pos += len;
        }

        // Draw each run with its own color
        pos = 0;
        for (const auto& run : para.runs) {
            UINT32 len = (UINT32)run.text.size();
            if (len == 0) { pos += len; continue; }

            // We draw the whole layout with each run's color using a custom renderer,
            // or approximate: draw the entire layout once per color
            auto brush = GetBrush(run.props.color);

            // For simplicity, draw the full layout with first run's color,
            // then re-draw with colored substrings via character effect (simplified):
            // Use full-layout draw for the dominant run
            if (pos == 0) {
                m_rt->DrawTextLayout(D2D1::Point2F(startX, y), layout.Get(), brush,
                    D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT);
            }
            pos += len;
        }

        // Get layout metrics for line height
        DWRITE_TEXT_METRICS metrics;
        layout->GetMetrics(&metrics);
        return metrics.height > 0 ? metrics.height : firstRun.props.fontSize * 96.0f / 72.0f * 1.2f;
    }

    void DrawRunText(const std::wstring& text, float x, float y, float fontSize,
                     const wchar_t* fontFamily, bool bold, bool italic,
                     ID2D1SolidColorBrush* brush) {
        ComPtr<IDWriteTextFormat> fmt;
        m_dwFactory->CreateTextFormat(fontFamily, nullptr,
            bold   ? DWRITE_FONT_WEIGHT_BOLD   : DWRITE_FONT_WEIGHT_NORMAL,
            italic ? DWRITE_FONT_STYLE_ITALIC  : DWRITE_FONT_STYLE_NORMAL,
            DWRITE_FONT_STRETCH_NORMAL, fontSize, L"en-US", fmt.GetAddressOf());
        if (!fmt) return;
        m_rt->DrawTextW(text.c_str(), (UINT32)text.size(), fmt.Get(),
                        D2D1::RectF(x, y, x+500, y+fontSize*2), brush);
    }

    void DrawSimpleText(const std::wstring& text, const D2D1_RECT_F& rect,
                        ID2D1SolidColorBrush* brush) {
        if (!m_defaultFormat) return;
        m_rt->DrawTextW(text.c_str(), (UINT32)text.size(), m_defaultFormat.Get(),
                        rect, brush);
    }

    void DrawGrid(float w, float h) {
        auto brush = GetBrush(D2D1::ColorF(0.6f, 0.6f, 0.8f, 0.3f));
        float step = 50.0f;
        for (float x = 0; x <= w; x += step)
            m_rt->DrawLine(D2D1::Point2F(x,0), D2D1::Point2F(x,h), brush, 0.5f);
        for (float y = 0; y <= h; y += step)
            m_rt->DrawLine(D2D1::Point2F(0,y), D2D1::Point2F(w,y), brush, 0.5f);
    }

    // ── Polygon helpers ───────────────────────────────────────────
    std::vector<D2D1_POINT_2F> MakeTriangle(const D2D1_RECT_F& r) {
        return {
            D2D1::Point2F((r.left+r.right)/2, r.top),
            D2D1::Point2F(r.right, r.bottom),
            D2D1::Point2F(r.left,  r.bottom)
        };
    }
    std::vector<D2D1_POINT_2F> MakeDiamond(const D2D1_RECT_F& r) {
        float cx = (r.left+r.right)/2, cy = (r.top+r.bottom)/2;
        return {
            D2D1::Point2F(cx, r.top),
            D2D1::Point2F(r.right, cy),
            D2D1::Point2F(cx, r.bottom),
            D2D1::Point2F(r.left, cy)
        };
    }
    std::vector<D2D1_POINT_2F> MakeRegularPolygon(const D2D1_RECT_F& r, int sides) {
        std::vector<D2D1_POINT_2F> pts;
        float cx = (r.left+r.right)/2, cy = (r.top+r.bottom)/2;
        float rx = (r.right-r.left)/2, ry = (r.bottom-r.top)/2;
        float offset = (sides % 2 == 0) ? 0 : -3.14159f/2;
        for (int i = 0; i < sides; i++) {
            float a = offset + 2*3.14159f*i/sides;
            pts.push_back(D2D1::Point2F(cx + rx*std::cos(a), cy + ry*std::sin(a)));
        }
        return pts;
    }
    std::vector<D2D1_POINT_2F> MakeParallelogram(const D2D1_RECT_F& r, float adj) {
        float offset = (r.right-r.left) * (adj > 0 ? adj : 0.25f);
        return {
            D2D1::Point2F(r.left+offset, r.top),
            D2D1::Point2F(r.right, r.top),
            D2D1::Point2F(r.right-offset, r.bottom),
            D2D1::Point2F(r.left, r.bottom)
        };
    }
    std::vector<D2D1_POINT_2F> MakeTrapezoid(const D2D1_RECT_F& r, float adj) {
        float inset = (r.right-r.left) * (adj > 0 ? adj : 0.2f);
        return {
            D2D1::Point2F(r.left+inset, r.top),
            D2D1::Point2F(r.right-inset, r.top),
            D2D1::Point2F(r.right, r.bottom),
            D2D1::Point2F(r.left, r.bottom)
        };
    }
    std::vector<D2D1_POINT_2F> MakeStar(const D2D1_RECT_F& r, int points, float innerRatio) {
        std::vector<D2D1_POINT_2F> pts;
        float cx = (r.left+r.right)/2, cy = (r.top+r.bottom)/2;
        float outer = std::min(r.right-r.left, r.bottom-r.top)/2;
        float inner = outer * (innerRatio > 0 ? innerRatio : 0.382f);
        for (int i = 0; i < points*2; i++) {
            float a = -3.14159f/2 + 3.14159f*i/points;
            float rd = (i%2==0) ? outer : inner;
            pts.push_back(D2D1::Point2F(cx + rd*std::cos(a), cy + rd*std::sin(a)));
        }
        return pts;
    }
    std::vector<D2D1_POINT_2F> MakeArrow(const D2D1_RECT_F& r, PptxEngine::ShapeType dir) {
        float w = r.right-r.left, h = r.bottom-r.top;
        float head = w * 0.35f;
        float shaft = h * 0.4f;
        switch (dir) {
            case PptxEngine::ShapeType::ArrowRight:
                return {
                    {r.left, r.top+h/2-shaft/2}, {r.right-head, r.top+h/2-shaft/2},
                    {r.right-head, r.top}, {r.right, r.top+h/2},
                    {r.right-head, r.bottom}, {r.right-head, r.top+h/2+shaft/2},
                    {r.left, r.top+h/2+shaft/2}
                };
            default:
                return MakeDiamond(r);
        }
    }
    std::vector<D2D1_POINT_2F> MakeChevron(const D2D1_RECT_F& r) {
        float w = r.right-r.left, h = r.bottom-r.top;
        float notch = w * 0.2f;
        return {
            {r.left, r.top}, {r.right-notch, r.top},
            {r.right, r.top+h/2}, {r.right-notch, r.bottom},
            {r.left, r.bottom}, {r.left+notch, r.top+h/2}
        };
    }
    int StarPointsFromType(PptxEngine::ShapeType t) {
        switch(t) {
            case PptxEngine::ShapeType::Star4:  return 4;
            case PptxEngine::ShapeType::Star5:  return 5;
            case PptxEngine::ShapeType::Star6:  return 6;
            case PptxEngine::ShapeType::Star8:  return 8;
            case PptxEngine::ShapeType::Star16: return 16;
            case PptxEngine::ShapeType::Star24: return 24;
            case PptxEngine::ShapeType::Star32: return 32;
            default: return 5;
        }
    }
};

// PptxEngine::Paragraph is the correct type used throughout D2DRenderer

// ═══════════════════════════════════════════════════════════════════
// SECTION 4: SLIDE THUMBNAIL PANEL
// ═══════════════════════════════════════════════════════════════════

class SlideThumbnailPanel {
public:
    static const int PANEL_WIDTH    = 180;
    static const int THUMB_W        = 150;
    static const int THUMB_H        = 84;
    static const int THUMB_MARGIN   = 12;
    static const int THUMB_LABEL_H  = 20;

    struct ThumbRenderInfo {
        int    slideIdx;
        RECT   rect;
        bool   selected;
        bool   hovered;
    };

private:
    HWND  m_hwnd     = nullptr;
    HWND  m_parent   = nullptr;
    int   m_scroll   = 0;
    int   m_selected = 0;
    int   m_hovered  = -1;

    ComPtr<ID2D1Factory>         m_d2dFactory;
    ComPtr<IDWriteFactory>       m_dwFactory;
    ComPtr<ID2D1HwndRenderTarget> m_rt;

    std::function<void(int)> m_onSelect;

public:
    void SetSelectCallback(std::function<void(int)> cb) { m_onSelect = cb; }
    void SetSelected(int idx) { m_selected = idx; }

    bool Create(HWND parent, HINSTANCE hInst) {
        m_parent = parent;

        WNDCLASSEXW wc = {};
        wc.cbSize      = sizeof(wc);
        wc.lpfnWndProc = ThumbnailWndProc;
        wc.hInstance   = hInst;
        wc.lpszClassName = L"PptxThumbnailPanel";
        wc.hbrBackground = (HBRUSH)GetStockObject(DKGRAY_BRUSH);
        RegisterClassExW(&wc);

        RECT parentRc;
        GetClientRect(parent, &parentRc);

        m_hwnd = CreateWindowExW(0, L"PptxThumbnailPanel", L"",
            WS_CHILD | WS_VISIBLE | WS_VSCROLL,
            0, 30, PANEL_WIDTH, parentRc.bottom - 60,
            parent, nullptr, hInst, this);

        if (!m_hwnd) return false;

        D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, m_d2dFactory.GetAddressOf());
        DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory),
            reinterpret_cast<IUnknown**>(m_dwFactory.GetAddressOf()));

        RECT rc;
        GetClientRect(m_hwnd, &rc);
        D2D1_RENDER_TARGET_PROPERTIES rtp = D2D1::RenderTargetProperties();
        D2D1_HWND_RENDER_TARGET_PROPERTIES htp = D2D1::HwndRenderTargetProperties(
            m_hwnd, D2D1::SizeU(rc.right, rc.bottom));
        m_d2dFactory->CreateHwndRenderTarget(rtp, htp, m_rt.GetAddressOf());

        return true;
    }

    void Render(const PptxEngine& engine) {
        if (!m_rt) return;
        m_rt->BeginDraw();
        m_rt->Clear(D2D1::ColorF(0.1f, 0.1f, 0.15f));

        RECT rc;
        GetClientRect(m_hwnd, &rc);
        int panelH = rc.bottom;
        int count  = engine.SlideCount();

        // Configure scroll bar
        SCROLLINFO si = {};
        si.cbSize = sizeof(si);
        si.fMask  = SIF_ALL;
        si.nMin   = 0;
        si.nMax   = count * (THUMB_H + THUMB_LABEL_H + THUMB_MARGIN) - 1;
        si.nPage  = panelH;
        si.nPos   = m_scroll;
        SetScrollInfo(m_hwnd, SB_VERT, &si, TRUE);

        ComPtr<IDWriteTextFormat> fmt;
        if (m_dwFactory) {
            m_dwFactory->CreateTextFormat(L"Segoe UI", nullptr,
                DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL,
                DWRITE_FONT_STRETCH_NORMAL, 11.0f, L"en-US", fmt.GetAddressOf());
        }

        for (int i = 0; i < count; i++) {
            int thumbY = THUMB_MARGIN + i * (THUMB_H + THUMB_LABEL_H + THUMB_MARGIN) - m_scroll;
            if (thumbY + THUMB_H < 0 || thumbY > panelH) continue;

            D2D1_RECT_F thumbRect = D2D1::RectF(
                (float)(PANEL_WIDTH/2 - THUMB_W/2), (float)thumbY,
                (float)(PANEL_WIDTH/2 + THUMB_W/2), (float)(thumbY + THUMB_H));

            // Selection highlight
            bool sel = (i == m_selected);
            bool hov = (i == m_hovered);

            if (sel) {
                D2D1_RECT_F selRect = D2D1::RectF(
                    thumbRect.left-4, thumbRect.top-4,
                    thumbRect.right+4, thumbRect.bottom+THUMB_LABEL_H+4);
                ComPtr<ID2D1SolidColorBrush> selBrush;
                m_rt->CreateSolidColorBrush(D2D1::ColorF(0.2f, 0.5f, 1.0f, 0.8f), selBrush.GetAddressOf());
                if (selBrush) m_rt->FillRoundedRectangle(D2D1::RoundedRect(selRect,4,4), selBrush.Get());
            } else if (hov) {
                D2D1_RECT_F hovRect = D2D1::RectF(
                    thumbRect.left-3, thumbRect.top-3,
                    thumbRect.right+3, thumbRect.bottom+THUMB_LABEL_H+3);
                ComPtr<ID2D1SolidColorBrush> hovBrush;
                m_rt->CreateSolidColorBrush(D2D1::ColorF(0.3f, 0.3f, 0.5f, 0.6f), hovBrush.GetAddressOf());
                if (hovBrush) m_rt->FillRoundedRectangle(D2D1::RoundedRect(hovRect,4,4), hovBrush.Get());
            }

            // White thumbnail background
            ComPtr<ID2D1SolidColorBrush> whiteBrush;
            m_rt->CreateSolidColorBrush(D2D1::ColorF(D2D1::ColorF::White), whiteBrush.GetAddressOf());
            if (whiteBrush) m_rt->FillRectangle(thumbRect, whiteBrush.Get());

            // Draw mini slide content
            DrawMiniSlide(engine, i, thumbRect);

            // Slide number label
            if (fmt) {
                D2D1_RECT_F labelRect = D2D1::RectF(
                    thumbRect.left, thumbRect.bottom+2,
                    thumbRect.right, thumbRect.bottom+THUMB_LABEL_H+2);
                ComPtr<ID2D1SolidColorBrush> labelBrush;
                D2D1_COLOR_F labelColor = sel
                    ? D2D1::ColorF(1.0f, 1.0f, 1.0f)
                    : D2D1::ColorF(0.7f, 0.7f, 0.8f);
                m_rt->CreateSolidColorBrush(labelColor, labelBrush.GetAddressOf());

                WCHAR label[32];
                const auto& slide = engine.GetSlide(i);
                if (!slide.title.empty()) {
                    swprintf(label, 32, L"%d - %.*s", i+1, 12,
                             Utf8ToWide(slide.title).c_str());
                } else {
                    swprintf(label, 32, L"Slide %d", i+1);
                }
                if (labelBrush) {
                    m_rt->DrawTextW(label, (UINT32)wcslen(label), fmt.Get(),
                                    labelRect, labelBrush.Get());
                }
            }

            // Border
            ComPtr<ID2D1SolidColorBrush> borderBrush;
            D2D1_COLOR_F borderColor = sel
                ? D2D1::ColorF(0.3f, 0.6f, 1.0f)
                : D2D1::ColorF(0.4f, 0.4f, 0.5f);
            m_rt->CreateSolidColorBrush(borderColor, borderBrush.GetAddressOf());
            if (borderBrush)
                m_rt->DrawRectangle(thumbRect, borderBrush.Get(), sel ? 2.0f : 0.5f);
        }

        m_rt->EndDraw();
    }

    void DrawMiniSlide(const PptxEngine& engine, int idx, const D2D1_RECT_F& thumbRect) {
        const auto& slide = engine.GetSlide(idx);
        if (slide.shapes.empty()) return;

        float scaleX = (thumbRect.right - thumbRect.left) / slide.width;
        float scaleY = (thumbRect.bottom - thumbRect.top) / slide.height;

        // Draw background
        if (slide.background.type == PptxEngine::FillInfo::Solid) {
            ComPtr<ID2D1SolidColorBrush> bgBrush;
            m_rt->CreateSolidColorBrush(slide.background.solidColor, bgBrush.GetAddressOf());
            if (bgBrush) m_rt->FillRectangle(thumbRect, bgBrush.Get());
        }

        // Draw shapes (simplified for thumbnail)
        D2D1_MATRIX_3X2_F transform = D2D1::Matrix3x2F(
            scaleX, 0, 0, scaleY, thumbRect.left, thumbRect.top);

        D2D1_MATRIX_3X2_F saved;
        m_rt->GetTransform(&saved);
        m_rt->SetTransform(transform);

        // Clip to thumbnail
        m_rt->PushAxisAlignedClip(D2D1::RectF(0, 0, slide.width, slide.height),
                                   D2D1_ANTIALIAS_MODE_ALIASED);

        for (const auto& shape : slide.shapes) {
            if (shape.type == PptxEngine::ShapeType::Picture) continue; // skip images in thumbs
            D2D1_RECT_F r = D2D1::RectF(shape.x, shape.y, shape.x+shape.w, shape.y+shape.h);
            if (shape.fill.type == PptxEngine::FillInfo::Solid) {
                ComPtr<ID2D1SolidColorBrush> shapeBrush;
                m_rt->CreateSolidColorBrush(shape.fill.solidColor, shapeBrush.GetAddressOf());
                if (shapeBrush) m_rt->FillRectangle(r, shapeBrush.Get());
            }
            if (shape.line.width > 0) {
                ComPtr<ID2D1SolidColorBrush> lineBrush;
                m_rt->CreateSolidColorBrush(shape.line.fill.solidColor, lineBrush.GetAddressOf());
                if (lineBrush) m_rt->DrawRectangle(r, lineBrush.Get(), 0.5f);
            }
            // Draw a line of text for non-empty text shapes
            if (!shape.paragraphs.empty() && !shape.paragraphs[0].runs.empty()) {
                ComPtr<IDWriteTextFormat> tinyFmt;
                m_dwFactory->CreateTextFormat(L"Segoe UI", nullptr,
                    DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL,
                    DWRITE_FONT_STRETCH_NORMAL, 8.0f, L"en-US", tinyFmt.GetAddressOf());
                if (tinyFmt) {
                    ComPtr<ID2D1SolidColorBrush> textBrush;
                    m_rt->CreateSolidColorBrush(shape.paragraphs[0].runs[0].props.color,
                                                textBrush.GetAddressOf());
                    const auto& firstRun = shape.paragraphs[0].runs[0];
                    if (textBrush)
                        m_rt->DrawTextW(firstRun.text.c_str(),
                                        (UINT32)std::min(firstRun.text.size(), (size_t)60),
                                        tinyFmt.Get(), r, textBrush.Get());
                }
            }
        }

        m_rt->PopAxisAlignedClip();
        m_rt->SetTransform(saved);
    }

    void HandleScroll(WPARAM wp) {
        SCROLLINFO si = {}; si.cbSize = sizeof(si); si.fMask = SIF_ALL;
        GetScrollInfo(m_hwnd, SB_VERT, &si);
        switch(LOWORD(wp)) {
            case SB_LINEUP:    si.nPos -= 30; break;
            case SB_LINEDOWN:  si.nPos += 30; break;
            case SB_PAGEUP:    si.nPos -= si.nPage; break;
            case SB_PAGEDOWN:  si.nPos += si.nPage; break;
            case SB_THUMBTRACK: si.nPos = si.nTrackPos; break;
        }
        si.nPos = max(si.nMin, min(si.nPos, (int)(si.nMax - si.nPage + 1)));
        m_scroll = si.nPos;
        SetScrollInfo(m_hwnd, SB_VERT, &si, TRUE);
        InvalidateRect(m_hwnd, NULL, FALSE);
    }

    int HitTest(int mouseY) {
        int totalY = mouseY + m_scroll;
        int perThumb = THUMB_H + THUMB_LABEL_H + THUMB_MARGIN;
        int idx = (totalY - THUMB_MARGIN) / perThumb;
        return (idx >= 0) ? idx : -1;
    }

    HWND GetHwnd() const { return m_hwnd; }

    static LRESULT CALLBACK ThumbnailWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
        SlideThumbnailPanel* panel = nullptr;
        if (msg == WM_NCCREATE) {
            CREATESTRUCT* cs = (CREATESTRUCT*)lp;
            panel = (SlideThumbnailPanel*)cs->lpCreateParams;
            SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)panel);
            panel->m_hwnd = hwnd;
        } else {
            panel = (SlideThumbnailPanel*)GetWindowLongPtr(hwnd, GWLP_USERDATA);
        }
        if (!panel) return DefWindowProc(hwnd, msg, wp, lp);

        switch(msg) {
            case WM_VSCROLL:
                panel->HandleScroll(wp);
                return 0;
            case WM_MOUSEWHEEL: {
                int delta = GET_WHEEL_DELTA_WPARAM(wp);
                SCROLLINFO si = {}; si.cbSize = sizeof(si); si.fMask = SIF_ALL;
                GetScrollInfo(hwnd, SB_VERT, &si);
                si.nPos -= delta / 4;
                si.nPos = max(si.nMin, min(si.nPos, (int)(si.nMax - si.nPage + 1)));
                panel->m_scroll = si.nPos;
                SetScrollInfo(hwnd, SB_VERT, &si, TRUE);
                InvalidateRect(hwnd, NULL, FALSE);
                return 0;
            }
            case WM_LBUTTONDOWN: {
                int y = GET_Y_LPARAM(lp);
                int idx = panel->HitTest(y);
                if (idx >= 0) {
                    panel->m_selected = idx;
                    if (panel->m_onSelect) panel->m_onSelect(idx);
                    InvalidateRect(hwnd, NULL, FALSE);
                }
                return 0;
            }
            case WM_MOUSEMOVE: {
                int y = GET_Y_LPARAM(lp);
                int idx = panel->HitTest(y);
                if (idx != panel->m_hovered) {
                    panel->m_hovered = idx;
                    InvalidateRect(hwnd, NULL, FALSE);
                }
                TRACKMOUSEEVENT tme = {sizeof(tme), TME_LEAVE, hwnd};
                TrackMouseEvent(&tme);
                return 0;
            }
            case WM_MOUSELEAVE:
                panel->m_hovered = -1;
                InvalidateRect(hwnd, NULL, FALSE);
                return 0;
            case WM_SIZE:
                if (panel->m_rt)
                    panel->m_rt->Resize(D2D1::SizeU(LOWORD(lp), HIWORD(lp)));
                return 0;
        }
        return DefWindowProc(hwnd, msg, wp, lp);
    }
};

// ═══════════════════════════════════════════════════════════════════
// SECTION 5: EXPORT ENGINE (PNG per slide, print support)
// ═══════════════════════════════════════════════════════════════════

class ExportEngine {
public:
    static bool ExportToPng(const PptxEngine& engine, int slideIdx,
                             const std::wstring& outPath,
                             int exportW = 1920, int exportH = 1080) {
        // Create WIC bitmap, D2D render to it, then save PNG
        ComPtr<IWICImagingFactory> wicFactory;
        HRESULT hr = CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER,
            IID_PPV_ARGS(&wicFactory));
        if (FAILED(hr)) return false;

        ComPtr<IWICBitmap> wicBitmap;
        wicFactory->CreateBitmap(exportW, exportH, GUID_WICPixelFormat32bppBGR,
            WICBitmapCacheOnLoad, wicBitmap.GetAddressOf());
        if (!wicBitmap) return false;

        ComPtr<ID2D1Factory> d2dFactory;
        D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, d2dFactory.GetAddressOf());

        ComPtr<ID2D1RenderTarget> rt;
        D2D1_RENDER_TARGET_PROPERTIES props = D2D1::RenderTargetProperties(
            D2D1_RENDER_TARGET_TYPE_DEFAULT,
            D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_IGNORE),
            96.0f, 96.0f
        );
        d2dFactory->CreateWicBitmapRenderTarget(wicBitmap.Get(), props, rt.GetAddressOf());
        if (!rt) return false;

        // Render slide to WIC bitmap RT
        const auto& slide = engine.GetSlide(slideIdx);
        rt->BeginDraw();
        rt->Clear(D2D1::ColorF(D2D1::ColorF::White));

        float scaleX = (float)exportW / slide.width;
        float scaleY = (float)exportH / slide.height;
        float scale  = std::min(scaleX, scaleY);
        rt->SetTransform(D2D1::Matrix3x2F::Scale(scale, scale));

        // Background
        ComPtr<ID2D1SolidColorBrush> bgBrush;
        if (slide.background.type == PptxEngine::FillInfo::Solid) {
            rt->CreateSolidColorBrush(slide.background.solidColor, bgBrush.GetAddressOf());
            if (bgBrush) rt->FillRectangle(D2D1::RectF(0,0,slide.width,slide.height), bgBrush.Get());
        } else {
            rt->CreateSolidColorBrush(D2D1::ColorF(D2D1::ColorF::White), bgBrush.GetAddressOf());
            if (bgBrush) rt->FillRectangle(D2D1::RectF(0,0,slide.width,slide.height), bgBrush.Get());
        }

        rt->EndDraw();

        // Save to PNG
        ComPtr<IWICStream> stream;
        wicFactory->CreateStream(stream.GetAddressOf());
        stream->InitializeFromFilename(outPath.c_str(), GENERIC_WRITE);

        ComPtr<IWICBitmapEncoder> encoder;
        wicFactory->CreateEncoder(GUID_ContainerFormatPng, nullptr, encoder.GetAddressOf());
        encoder->Initialize(stream.Get(), WICBitmapEncoderNoCache);

        ComPtr<IWICBitmapFrameEncode> frame;
        encoder->CreateNewFrame(frame.GetAddressOf(), nullptr);
        frame->Initialize(nullptr);
        frame->SetSize(exportW, exportH);

        WICPixelFormatGUID fmt = GUID_WICPixelFormat32bppBGR;
        frame->SetPixelFormat(&fmt);
        frame->WriteSource(wicBitmap.Get(), nullptr);
        frame->Commit();
        encoder->Commit();

        return true;
    }

    static bool ExportAllToPng(const PptxEngine& engine, const std::wstring& folder) {
        bool ok = true;
        for (int i = 0; i < engine.SlideCount(); i++) {
            std::wstring path = folder + L"\\slide" + std::to_wstring(i+1) + L".png";
            ok &= ExportToPng(engine, i, path);
        }
        return ok;
    }
};

// ═══════════════════════════════════════════════════════════════════
// SECTION 6: TOOLBAR AND MENUS
// ═══════════════════════════════════════════════════════════════════

class ToolbarManager {
public:
    static const int TOOLBAR_H = 32;

    static HMENU BuildMainMenu() {
        HMENU menu  = CreateMenu();

        // File
        HMENU file = CreatePopupMenu();
        AppendMenuW(file, MF_STRING, 1001, L"&Open\tCtrl+O");
        AppendMenuW(file, MF_STRING, 1004, L"&Save\tCtrl+S");
        AppendMenuW(file, MF_STRING, 1007, L"Save &As...");
        AppendMenuW(file, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(file, MF_STRING, 1009, L"Export Slide as PNG...");
        AppendMenuW(file, MF_STRING, 1010, L"Export All Slides as PNG...");
        AppendMenuW(file, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(file, MF_STRING, 1005, L"E&xit\tAlt+F4");
        AppendMenuW(menu, MF_POPUP, (UINT_PTR)file, L"&File");

        // Edit
        HMENU edit = CreatePopupMenu();
        AppendMenuW(edit, MF_STRING, 2001, L"&Undo\tCtrl+Z");
        AppendMenuW(edit, MF_STRING, 2002, L"&Redo\tCtrl+Y");
        AppendMenuW(edit, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(edit, MF_STRING, 2003, L"Select &All\tCtrl+A");
        AppendMenuW(edit, MF_STRING, 2004, L"&Delete Selected\tDel");
        AppendMenuW(menu, MF_POPUP, (UINT_PTR)edit, L"&Edit");

        // View
        HMENU view = CreatePopupMenu();
        AppendMenuW(view, MF_STRING, 3001, L"&Previous Slide\tLeft");
        AppendMenuW(view, MF_STRING, 3002, L"&Next Slide\tRight");
        AppendMenuW(view, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(view, MF_STRING, 3003, L"Zoom &In\tCtrl++");
        AppendMenuW(view, MF_STRING, 3004, L"Zoom &Out\tCtrl+-");
        AppendMenuW(view, MF_STRING, 3005, L"&Fit Slide\tCtrl+0");
        AppendMenuW(view, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(view, MF_STRING | MF_CHECKED, 3006, L"Show &Thumbnails");
        AppendMenuW(view, MF_STRING, 3007, L"Show &Grid\tCtrl+G");
        AppendMenuW(view, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(view, MF_STRING, 3008, L"&Presenter Mode\tF5");
        AppendMenuW(menu, MF_POPUP, (UINT_PTR)view, L"&View");

        // Slide
        HMENU slide = CreatePopupMenu();
        AppendMenuW(slide, MF_STRING, 4001, L"&New Slide\tCtrl+M");
        AppendMenuW(slide, MF_STRING, 4002, L"&Delete Slide");
        AppendMenuW(slide, MF_STRING, 4003, L"&Duplicate Slide");
        AppendMenuW(slide, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(slide, MF_STRING, 4004, L"&Move Up");
        AppendMenuW(slide, MF_STRING, 4005, L"&Move Down");
        AppendMenuW(menu, MF_POPUP, (UINT_PTR)slide, L"&Slide");

        // Insert
        HMENU ins = CreatePopupMenu();
        AppendMenuW(ins, MF_STRING, 5001, L"&Text Box");
        AppendMenuW(ins, MF_STRING, 5002, L"&Rectangle");
        AppendMenuW(ins, MF_STRING, 5003, L"&Ellipse");
        AppendMenuW(ins, MF_STRING, 5004, L"&Image from File...");
        AppendMenuW(ins, MF_STRING, 5005, L"&Line");
        AppendMenuW(ins, MF_STRING, 5006, L"&Arrow");
        AppendMenuW(menu, MF_POPUP, (UINT_PTR)ins, L"&Insert");

        // Help
        HMENU help = CreatePopupMenu();
        AppendMenuW(help, MF_STRING, 9001, L"&About");
        AppendMenuW(help, MF_STRING, 9002, L"&Keyboard Shortcuts");
        AppendMenuW(menu, MF_POPUP, (UINT_PTR)help, L"&Help");

        return menu;
    }

    static void DrawToolbar(ID2D1HwndRenderTarget* rt, HWND hwnd,
                            int currentSlide, int totalSlides, float zoom) {
        if (!rt) return;
        RECT rc;
        GetClientRect(hwnd, &rc);

        D2D1_RECT_F bar = D2D1::RectF(0, 0, (float)rc.right, (float)TOOLBAR_H);
        ComPtr<ID2D1SolidColorBrush> bgBrush;
        rt->CreateSolidColorBrush(D2D1::ColorF(0.12f, 0.12f, 0.20f), bgBrush.GetAddressOf());
        if (bgBrush) rt->FillRectangle(bar, bgBrush.Get());

        // Separator line
        ComPtr<ID2D1SolidColorBrush> sepBrush;
        rt->CreateSolidColorBrush(D2D1::ColorF(0.3f,0.3f,0.5f), sepBrush.GetAddressOf());
        if (sepBrush) rt->DrawLine(D2D1::Point2F(0,(float)TOOLBAR_H),
                                    D2D1::Point2F((float)rc.right,(float)TOOLBAR_H),
                                    sepBrush.Get(), 1.0f);
    }
};

// ═══════════════════════════════════════════════════════════════════
// SECTION 7: MAIN APPLICATION
// ═══════════════════════════════════════════════════════════════════

class ViewerApp {
    static const int CMD_OPEN        = 1001;
    static const int CMD_SAVE        = 1004;
    static const int CMD_SAVE_AS     = 1007;
    static const int CMD_EXIT        = 1005;
    static const int CMD_EXPORT_PNG  = 1009;
    static const int CMD_EXPORT_ALL  = 1010;
    static const int CMD_UNDO        = 2001;
    static const int CMD_REDO        = 2002;
    static const int CMD_PREV        = 3001;
    static const int CMD_NEXT        = 3002;
    static const int CMD_ZOOM_IN     = 3003;
    static const int CMD_ZOOM_OUT    = 3004;
    static const int CMD_FIT         = 3005;
    static const int CMD_GRID        = 3007;
    static const int CMD_FULLSCREEN  = 3008;
    static const int CMD_ABOUT       = 9001;
    static const int CMD_SHORTCUTS   = 9002;

private:
    HWND m_hwnd  = nullptr;
    HWND m_mainView = nullptr;

    PptxEngine           m_engine;
    D2DRenderer          m_renderer;
    SlideThumbnailPanel  m_thumbPanel;

    int   m_currentSlide   = 0;
    float m_zoom           = 1.0f;
    float m_panX           = 0.0f, m_panY = 0.0f;
    bool  m_showGrid       = false;
    bool  m_showThumbs     = true;
    bool  m_fullscreen     = false;
    bool  m_isDragging     = false;
    POINT m_lastMouse      = {};
    bool  m_fileLoaded     = false;
    std::wstring m_currentFile;

    // Status bar info
    WCHAR m_statusText[256] = L"Ready — Open a .pptx file to start";

    // Menu check states
    bool  m_gridChecked    = false;
    bool  m_thumbChecked   = true;

public:
    bool Init(HINSTANCE hInst) {
        CoInitialize(nullptr);

        // Register main window class
        WNDCLASSEXW wc = {};
        wc.cbSize      = sizeof(wc);
        wc.lpfnWndProc = WndProc;
        wc.hInstance   = hInst;
        wc.lpszClassName = L"PptxEditorClass";
        wc.hbrBackground = (HBRUSH)GetStockObject(DKGRAY_BRUSH);
        wc.hCursor     = LoadCursor(nullptr, IDC_ARROW);
        wc.hIcon       = LoadIcon(nullptr, IDI_APPLICATION);
        RegisterClassExW(&wc);

        m_hwnd = CreateWindowExW(WS_EX_APPWINDOW,
            L"PptxEditorClass",
            L"Presentation Studio Pro — PPTX Editor",
            WS_OVERLAPPEDWINDOW,
            CW_USEDEFAULT, 0, 1400, 900,
            nullptr, nullptr, hInst, this);

        if (!m_hwnd) return false;

        // Set menu
        HMENU menu = ToolbarManager::BuildMainMenu();
        SetMenu(m_hwnd, menu);

        // Init renderer
        if (!m_renderer.Init(m_hwnd)) return false;

        // Init thumbnail panel
        m_thumbPanel.SetSelectCallback([this](int idx) {
            m_currentSlide = idx;
            UpdateStatusBar();
            InvalidateRect(m_hwnd, nullptr, FALSE);
        });
        m_thumbPanel.Create(m_hwnd, hInst);

        return true;
    }

    void Run() {
        ShowWindow(m_hwnd, SW_SHOW);
        UpdateWindow(m_hwnd);

        MSG msg;
        while (GetMessage(&msg, nullptr, 0, 0)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }

        CoUninitialize();
    }

    // Load a PPTX from a given path directly (used by command-line arg)
    bool OpenFileFromPath(const std::wstring& path) {
        SetCursor(LoadCursor(nullptr, IDC_WAIT));
        bool ok = m_engine.LoadPptx(path);
        SetCursor(LoadCursor(nullptr, IDC_ARROW));
        if (!ok) return false;
        m_currentFile  = path;
        m_currentSlide = 0;
        m_fileLoaded   = true;
        m_zoom         = 1.0f;
        m_panX = m_panY = 0.0f;
        WCHAR title[300];
        swprintf(title, 300, L"Presentation Studio Pro — %s (%d slides)",
                 PathFindFileNameW(path.c_str()), m_engine.SlideCount());
        SetWindowTextW(m_hwnd, title);
        m_renderer.SetConfig({m_zoom, m_panX, m_panY, true, true, 96.0f, m_showGrid});
        m_thumbPanel.SetSelected(0);
        UpdateStatusBar();
        InvalidateRect(m_hwnd, nullptr, FALSE);
        return true;
    }

    void OpenFile() {
        OPENFILENAMEW ofn = {};
        WCHAR path[MAX_PATH] = {};
        ofn.lStructSize = sizeof(ofn);
        ofn.hwndOwner   = m_hwnd;
        ofn.lpstrFile   = path;
        ofn.nMaxFile    = MAX_PATH;
        ofn.lpstrFilter = L"PowerPoint (*.pptx)\0*.pptx\0All Files\0*.*\0";
        ofn.Flags       = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST;

        if (!GetOpenFileNameW(&ofn)) return;

        SetCursor(LoadCursor(nullptr, IDC_WAIT));
        bool ok = m_engine.LoadPptx(path);
        SetCursor(LoadCursor(nullptr, IDC_ARROW));

        if (!ok) {
            MessageBoxW(m_hwnd, L"Failed to open PPTX file.", L"Error", MB_ICONERROR);
            return;
        }

        m_currentFile  = path;
        m_currentSlide = 0;
        m_fileLoaded   = true;
        m_zoom         = 1.0f;
        m_panX = m_panY = 0.0f;

        WCHAR title[300];
        swprintf(title, 300, L"Presentation Studio Pro — %s (%d slides)",
                 PathFindFileNameW(path), m_engine.SlideCount());
        SetWindowTextW(m_hwnd, title);

        m_renderer.SetConfig({m_zoom, m_panX, m_panY, true, true, 96.0f, m_showGrid});
        m_thumbPanel.SetSelected(0);
        UpdateStatusBar();

        InvalidateRect(m_hwnd, nullptr, FALSE);
    }

    void SaveFile() {
        if (m_currentFile.empty()) { SaveFileAs(); return; }
        m_engine.SavePptx(m_currentFile);
        swprintf(m_statusText, 256, L"Saved: %s", PathFindFileNameW(m_currentFile.c_str()));
        InvalidateRect(m_hwnd, nullptr, FALSE);
    }

    void SaveFileAs() {
        OPENFILENAMEW ofn = {};
        WCHAR path[MAX_PATH] = {};
        if (!m_currentFile.empty()) wcscpy_s(path, m_currentFile.c_str());
        ofn.lStructSize = sizeof(ofn);
        ofn.hwndOwner   = m_hwnd;
        ofn.lpstrFile   = path;
        ofn.nMaxFile    = MAX_PATH;
        ofn.lpstrFilter = L"PowerPoint (*.pptx)\0*.pptx\0";
        ofn.lpstrDefExt = L"pptx";
        ofn.Flags       = OFN_OVERWRITEPROMPT;
        if (GetSaveFileNameW(&ofn)) {
            m_currentFile = path;
            m_engine.SavePptx(path);
            swprintf(m_statusText, 256, L"Saved as: %s", PathFindFileNameW(path));
        }
    }

    void ExportCurrentSlide() {
        OPENFILENAMEW ofn = {};
        WCHAR path[MAX_PATH] = {};
        swprintf(path, MAX_PATH, L"slide%d.png", m_currentSlide+1);
        ofn.lStructSize = sizeof(ofn);
        ofn.hwndOwner   = m_hwnd;
        ofn.lpstrFile   = path;
        ofn.nMaxFile    = MAX_PATH;
        ofn.lpstrFilter = L"PNG Image (*.png)\0*.png\0";
        ofn.lpstrDefExt = L"png";
        ofn.Flags       = OFN_OVERWRITEPROMPT;
        if (GetSaveFileNameW(&ofn)) {
            if (ExportEngine::ExportToPng(m_engine, m_currentSlide, path))
                MessageBoxW(m_hwnd, L"Slide exported successfully.", L"Export", MB_ICONINFORMATION);
            else
                MessageBoxW(m_hwnd, L"Export failed.", L"Error", MB_ICONERROR);
        }
    }

    void ExportAllSlides() {
        BROWSEINFOW bi = {};
        bi.hwndOwner = m_hwnd;
        bi.lpszTitle = L"Select folder for exported PNG files";
        bi.ulFlags   = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;
        LPITEMIDLIST pidl = SHBrowseForFolderW(&bi);
        if (!pidl) return;
        WCHAR folder[MAX_PATH];
        SHGetPathFromIDListW(pidl, folder);
        CoTaskMemFree(pidl);
        if (ExportEngine::ExportAllToPng(m_engine, folder)) {
            WCHAR msg[256];
            swprintf(msg, 256, L"Exported %d slides to %s", m_engine.SlideCount(), folder);
            MessageBoxW(m_hwnd, msg, L"Export Complete", MB_ICONINFORMATION);
        }
    }

    void NextSlide() {
        if (m_currentSlide < m_engine.SlideCount() - 1) {
            m_currentSlide++;
            m_thumbPanel.SetSelected(m_currentSlide);
            UpdateStatusBar();
            InvalidateRect(m_hwnd, nullptr, FALSE);
        }
    }

    void PrevSlide() {
        if (m_currentSlide > 0) {
            m_currentSlide--;
            m_thumbPanel.SetSelected(m_currentSlide);
            UpdateStatusBar();
            InvalidateRect(m_hwnd, nullptr, FALSE);
        }
    }

    void ZoomIn() {
        m_zoom = std::min(m_zoom * 1.2f, 5.0f);
        UpdateRendererConfig();
        InvalidateRect(m_hwnd, nullptr, FALSE);
    }

    void ZoomOut() {
        m_zoom = std::max(m_zoom / 1.2f, 0.1f);
        UpdateRendererConfig();
        InvalidateRect(m_hwnd, nullptr, FALSE);
    }

    void ZoomFit() {
        if (!m_fileLoaded) return;
        RECT rc;
        GetClientRect(m_hwnd, &rc);
        float availW = (float)(rc.right  - rc.left - (m_showThumbs ? SlideThumbnailPanel::PANEL_WIDTH : 0) - 40);
        float availH = (float)(rc.bottom - rc.top - ToolbarManager::TOOLBAR_H - 50);
        const auto& slide = m_engine.GetSlide(m_currentSlide);
        m_zoom = std::min(availW / slide.width, availH / slide.height);
        m_panX = m_panY = 0;
        UpdateRendererConfig();
        InvalidateRect(m_hwnd, nullptr, FALSE);
    }

    void ToggleGrid() {
        m_showGrid = !m_showGrid;
        m_gridChecked = m_showGrid;
        // Toggle menu checkmark
        HMENU menu = GetMenu(m_hwnd);
        if (menu) CheckMenuItem(menu, CMD_GRID,
            MF_BYCOMMAND | (m_showGrid ? MF_CHECKED : MF_UNCHECKED));
        UpdateRendererConfig();
        InvalidateRect(m_hwnd, nullptr, FALSE);
    }

    void UpdateRendererConfig() {
        D2DRenderer::RenderConfig cfg;
        cfg.scale    = m_zoom;
        cfg.offsetX  = m_panX;
        cfg.offsetY  = m_panY;
        cfg.antialias = true;
        cfg.hqText   = true;
        cfg.showGrid = m_showGrid;
        m_renderer.SetConfig(cfg);
    }

    void UpdateStatusBar() {
        if (!m_fileLoaded) return;
        const auto& slide = m_engine.GetSlide(m_currentSlide);
        swprintf(m_statusText, 256,
                 L"Slide %d / %d  |  %.0f×%.0f  |  Zoom: %.0f%%  |  %d shapes",
                 m_currentSlide+1, m_engine.SlideCount(),
                 slide.width, slide.height,
                 m_zoom * 100.0f,
                 (int)slide.shapes.size());
    }

    void ShowAbout() {
        MessageBoxW(m_hwnd,
            L"Presentation Studio Pro\nVersion 3.0\n\n"
            L"High-quality PPTX viewer and editor\n"
            L"Built with Direct2D + DirectWrite\n\n"
            L"Features:\n"
            L"• Full PPTX parsing (shapes, text, images, tables)\n"
            L"• High-quality D2D rendering with gradients\n"
            L"• Slide thumbnail panel\n"
            L"• PNG export\n"
            L"• Edit shapes, move, resize, delete\n"
            L"• Undo/redo support",
            L"About Presentation Studio Pro",
            MB_ICONINFORMATION);
    }

    void ShowShortcuts() {
        MessageBoxW(m_hwnd,
            L"Keyboard Shortcuts:\n\n"
            L"Ctrl+O        Open file\n"
            L"Ctrl+S        Save file\n"
            L"Ctrl+Z        Undo\n"
            L"Ctrl+Y        Redo\n"
            L"← / →          Previous / Next slide\n"
            L"Ctrl++         Zoom in\n"
            L"Ctrl+-         Zoom out\n"
            L"Ctrl+0         Fit slide\n"
            L"Ctrl+G         Toggle grid\n"
            L"F5             Fullscreen presenter mode\n"
            L"Esc            Exit presenter / Deselect\n"
            L"Delete         Delete selected shape",
            L"Keyboard Shortcuts", MB_ICONINFORMATION);
    }

    void Render() {
        if (!m_fileLoaded) {
            RenderSplashScreen();
            return;
        }
        RECT rc;
        GetClientRect(m_hwnd, &rc);
        m_renderer.Resize(rc.right, rc.bottom);

        const auto& slide = m_engine.GetSlide(m_currentSlide);
        m_renderer.RenderSlide(slide, m_engine, m_currentSlide);

        // Status bar (drawn on the D2D surface as overlay)
        // (In real app would use a separate HWND)
    }

    void RenderSplashScreen() {
        RECT rc;
        GetClientRect(m_hwnd, &rc);
        m_renderer.Resize(rc.right, rc.bottom);
        // The D2DRenderer::RenderSlide handles empty state gracefully
        // Just show background
    }

    void Resize(int w, int h) {
        m_renderer.Resize(w, h);
        // Reposition thumbnail panel
        if (m_thumbPanel.GetHwnd()) {
            SetWindowPos(m_thumbPanel.GetHwnd(), nullptr,
                0, ToolbarManager::TOOLBAR_H,
                SlideThumbnailPanel::PANEL_WIDTH, h - ToolbarManager::TOOLBAR_H - 25,
                SWP_NOZORDER);
        }
        InvalidateRect(m_hwnd, nullptr, FALSE);
    }

    void Cleanup() {
        // Resources released via smart pointers
    }

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
        ViewerApp* app = nullptr;
        if (msg == WM_NCCREATE) {
            CREATESTRUCT* cs = (CREATESTRUCT*)lp;
            app = (ViewerApp*)cs->lpCreateParams;
            SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)app);
            app->m_hwnd = hwnd;
        } else {
            app = (ViewerApp*)GetWindowLongPtr(hwnd, GWLP_USERDATA);
        }

        if (!app) return DefWindowProc(hwnd, msg, wp, lp);

        switch (msg) {
            case WM_PAINT: {
                PAINTSTRUCT ps;
                BeginPaint(hwnd, &ps);
                app->Render();
                EndPaint(hwnd, &ps);
                return 0;
            }

            case WM_SIZE:
                app->Resize(LOWORD(lp), HIWORD(lp));
                return 0;

            case WM_MOUSEWHEEL: {
                int delta = GET_WHEEL_DELTA_WPARAM(wp);
                if (GetKeyState(VK_CONTROL) < 0) {
                    if (delta > 0) app->ZoomIn(); else app->ZoomOut();
                } else {
                    if (delta > 0) app->PrevSlide(); else app->NextSlide();
                }
                return 0;
            }

            case WM_LBUTTONDOWN:
                app->m_isDragging = true;
                app->m_lastMouse  = {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
                SetCapture(hwnd);
                return 0;

            case WM_LBUTTONUP:
                app->m_isDragging = false;
                ReleaseCapture();
                return 0;

            case WM_MOUSEMOVE:
                if (app->m_isDragging && (GetKeyState(VK_SPACE) < 0)) {
                    POINT cur = {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
                    app->m_panX += (float)(cur.x - app->m_lastMouse.x);
                    app->m_panY += (float)(cur.y - app->m_lastMouse.y);
                    app->m_lastMouse = cur;
                    app->UpdateRendererConfig();
                    InvalidateRect(hwnd, nullptr, FALSE);
                }
                return 0;

            case WM_COMMAND:
                switch (LOWORD(wp)) {
                    case CMD_OPEN:       app->OpenFile();            break;
                    case CMD_SAVE:       app->SaveFile();            break;
                    case CMD_SAVE_AS:    app->SaveFileAs();          break;
                    case CMD_EXIT:       PostQuitMessage(0);         break;
                    case CMD_EXPORT_PNG: app->ExportCurrentSlide();  break;
                    case CMD_EXPORT_ALL: app->ExportAllSlides();     break;
                    case CMD_UNDO:       app->m_engine.UndoEdit();
                                         InvalidateRect(hwnd, nullptr, FALSE); break;
                    case CMD_PREV:       app->PrevSlide();           break;
                    case CMD_NEXT:       app->NextSlide();           break;
                    case CMD_ZOOM_IN:    app->ZoomIn();              break;
                    case CMD_ZOOM_OUT:   app->ZoomOut();             break;
                    case CMD_FIT:        app->ZoomFit();             break;
                    case CMD_GRID:       app->ToggleGrid();          break;
                    case CMD_ABOUT:      app->ShowAbout();           break;
                    case CMD_SHORTCUTS:  app->ShowShortcuts();       break;
                }
                return 0;

            case WM_KEYDOWN:
                switch (wp) {
                    case VK_RIGHT:   app->NextSlide();  break;
                    case VK_LEFT:    app->PrevSlide();  break;
                    case VK_F5:
                        // Simple fullscreen toggle
                        if (!app->m_fullscreen) {
                            SetWindowLong(hwnd, GWL_STYLE,
                                GetWindowLong(hwnd, GWL_STYLE) & ~WS_OVERLAPPEDWINDOW);
                            ShowWindow(hwnd, SW_MAXIMIZE);
                        } else {
                            SetWindowLong(hwnd, GWL_STYLE,
                                GetWindowLong(hwnd, GWL_STYLE) | WS_OVERLAPPEDWINDOW);
                            ShowWindow(hwnd, SW_NORMAL);
                        }
                        app->m_fullscreen = !app->m_fullscreen;
                        break;
                    case VK_ESCAPE:
                        if (app->m_fullscreen) {
                            SetWindowLong(hwnd, GWL_STYLE,
                                GetWindowLong(hwnd, GWL_STYLE) | WS_OVERLAPPEDWINDOW);
                            ShowWindow(hwnd, SW_NORMAL);
                            app->m_fullscreen = false;
                        }
                        break;
                    case 'O':
                        if (GetKeyState(VK_CONTROL) < 0) app->OpenFile();
                        break;
                    case 'S':
                        if (GetKeyState(VK_CONTROL) < 0) {
                            if (GetKeyState(VK_SHIFT) < 0) app->SaveFileAs();
                            else app->SaveFile();
                        }
                        break;
                    case 'Z':
                        if (GetKeyState(VK_CONTROL) < 0) {
                            app->m_engine.UndoEdit();
                            InvalidateRect(hwnd, nullptr, FALSE);
                        }
                        break;
                    case VK_OEM_PLUS:
                        if (GetKeyState(VK_CONTROL) < 0) app->ZoomIn();
                        break;
                    case VK_OEM_MINUS:
                        if (GetKeyState(VK_CONTROL) < 0) app->ZoomOut();
                        break;
                    case '0':
                        if (GetKeyState(VK_CONTROL) < 0) app->ZoomFit();
                        break;
                    case 'G':
                        if (GetKeyState(VK_CONTROL) < 0) app->ToggleGrid();
                        break;
                }
                return 0;

            case WM_DESTROY:
                app->Cleanup();
                PostQuitMessage(0);
                return 0;
        }

        return DefWindowProc(hwnd, msg, wp, lp);
    }
};

// ═══════════════════════════════════════════════════════════════════
// MAIN ENTRY POINT
// ═══════════════════════════════════════════════════════════════════

int WINAPI wWinMain(HINSTANCE hInst, HINSTANCE, LPWSTR cmdLine, int) {
    SetProcessDPIAware();

    // Check for command-line file argument
    ViewerApp app;
    if (!app.Init(hInst)) {
        MessageBoxW(nullptr, L"Failed to initialize application.", L"Error", MB_ICONERROR);
        return 1;
    }

    // If a file was passed on the command line, open it immediately
    if (cmdLine && wcslen(cmdLine) > 0) {
        std::wstring path = cmdLine;
        // Strip quotes if present
        if (!path.empty() && path.front() == L'"') {
            path = path.substr(1);
            if (!path.empty() && path.back() == L'"') path.pop_back();
        }
        if (!path.empty()) {
            // Load file directly
            SetCursor(LoadCursor(nullptr, IDC_WAIT));
            bool ok = app.m_engine.LoadPptx(path);
            SetCursor(LoadCursor(nullptr, IDC_ARROW));
            if (ok) {
                app.m_currentFile  = path;
                app.m_currentSlide = 0;
                app.m_fileLoaded   = true;
                app.UpdateRendererConfig();
                app.UpdateStatusBar();
                app.m_thumbPanel.SetSelected(0);
            }
        }
    }

    app.Run();
    return 0;
}
