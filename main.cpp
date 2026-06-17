// pptx_editor.cpp
// Professional PowerPoint Viewer & Editor with Ribbon UI
// Compile (MinGW):
//   g++ -std=c++17 -O2 -o pptx_editor.exe main.cpp -lgdiplus -lole32 -loleaut32 -luuid -lshlwapi -lcomctl32 -mwindows
// Compile (MSVC):
//   cl /std:c++17 /O2 /EHsc main.cpp gdiplus.lib ole32.lib oleaut32.lib uuid.lib shlwapi.lib comctl32.lib /link /SUBSYSTEM:WINDOWS

#define _USE_MATH_DEFINES
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <gdiplus.h>
#include <objbase.h>
#include <shlwapi.h>
#include <commdlg.h>
#include <richedit.h>
#include <cmath>
#include <cstdlib>
#include <windowsx.h>
#include <string>
#include <vector>
#include <map>
#include <algorithm>
#include <fstream>
#include <sstream>

#pragma comment(lib, "gdiplus.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "comctl32.lib")

using namespace Gdiplus;
using namespace std;

// ==================== GLOBAL VARIABLES ====================
float g_scaleFactor = 1.0f;
HWND g_hwnd = NULL;
int g_windowWidth = 1600;
int g_windowHeight = 900;

// ==================== DATA STRUCTURES ====================
enum ShapeType {
    SHAPE_RECTANGLE = 0,
    SHAPE_ROUNDED_RECT = 1,
    SHAPE_ELLIPSE = 2,
    SHAPE_TRIANGLE = 3,
    SHAPE_DIAMOND = 4,
    SHAPE_ARROW_RIGHT = 5,
    SHAPE_STAR = 6
};

enum TextAlignment {
    ALIGN_LEFT = 0,
    ALIGN_CENTER = 1,
    ALIGN_RIGHT = 2
};

struct SlideElement {
    RECT rect;                    // Position and size (in percentage of slide)
    wstring text;
    COLORREF bgColor;
    COLORREF textColor;
    COLORREF borderColor;
    int borderWidth;
    float fontSize;
    bool isBold;
    bool isItalic;
    bool isUnderline;
    ShapeType shapeType;
    TextAlignment alignment;
    int id;
    bool isSelected;
    
    SlideElement() : bgColor(RGB(255, 255, 255)), textColor(RGB(0, 0, 0)), 
                     borderColor(RGB(0, 0, 0)), borderWidth(0), fontSize(18.0f),
                     isBold(false), isItalic(false), isUnderline(false),
                     shapeType(SHAPE_RECTANGLE), alignment(ALIGN_LEFT),
                     id(0), isSelected(false) {
        rect = {10, 10, 50, 30};
    }
};

struct Slide {
    vector<SlideElement> elements;
    wstring title;
    COLORREF bgColor;
    int id;
    bool isSelected;
    
    Slide() : bgColor(RGB(255, 255, 255)), id(0), isSelected(false) {
        title = L"Slide";
    }
};

// Ribbon tab
struct RibbonButton {
    wstring text;
    wstring icon;
    RECT rect;
    bool isHovered;
    bool isActive;
    int id;
    
    RibbonButton() : isHovered(false), isActive(false), id(0) {
        rect = {0, 0, 0, 0};
    }
};

struct RibbonGroup {
    wstring name;
    vector<RibbonButton> buttons;
    RECT rect;
};

struct RibbonTab {
    wstring name;
    vector<RibbonGroup> groups;
    bool isActive;
    RECT tabRect;
};

// ==================== APPLICATION STATE ====================
struct AppState {
    vector<Slide> slides;
    int currentSlide;
    int selectedElementId;
    bool isMaximized;
    bool isDragging;
    POINT dragStart;
    POINT lastMousePos;
    
    // UI panels
    int ribbonHeight;
    int sidebarWidth;
    int statusBarHeight;
    
    // Ribbon
    vector<RibbonTab> ribbonTabs;
    int activeTabIndex;
    
    // Button hover states
    map<int, bool> buttonHovered;
    
    // Slide thumbnails scroll
    int thumbnailScrollPos;
    
    // Editing mode
    bool isEditing;
    bool isResizing;
    int resizeHandle; // 0=none, 1-8=handles
    POINT resizeStart;
    RECT originalRect;
    
    // Color picker
    bool colorPickerOpen;
    int colorPickerTarget; // 0=fill, 1=text, 2=border
    
    // File path
    wstring currentFile;
    bool isModified;
    
    // Undo/Redo (simplified)
    vector<vector<Slide>> undoStack;
    int undoIndex;
    
    AppState() : currentSlide(0), selectedElementId(-1), isMaximized(false),
                 isDragging(false), dragStart{0, 0}, lastMousePos{0, 0},
                 ribbonHeight(120), sidebarWidth(200), statusBarHeight(25),
                 activeTabIndex(0), thumbnailScrollPos(0),
                 isEditing(false), isResizing(false), resizeHandle(0),
                 colorPickerOpen(false), colorPickerTarget(0),
                 isModified(false), undoIndex(-1) {
    }
};

AppState g_state;

// ==================== FORWARD DECLARATIONS ====================
void InitializeRibbon();
void InitializeDefaultSlides();
void SaveState();
void Undo();
void Redo();
void DrawRibbon(Graphics& g, int width);
void DrawSidebar(Graphics& g, int width, int height);
void DrawSlide(Graphics& g, int x, int y, int width, int height, const Slide& slide, bool isThumbnail);
void DrawStatusBar(Graphics& g, int width, int height);
void DrawColorPicker(Graphics& g, int x, int y);
void HandleRibbonClick(int x, int y);
void HandleSlideClick(int x, int y, int slideX, int slideY, int slideWidth, int slideHeight);
void HandleThumbnailClick(int x, int y, int sidebarWidth, int height);
void AddNewSlide();
void DeleteCurrentSlide();
void AddTextBox();
void AddShape(ShapeType type);
void ChangeFillColor(COLORREF color);
void ChangeTextColor(COLORREF color);
void ChangeBorderColor(COLORREF color);
void ToggleBold();
void ToggleItalic();
void ToggleUnderline();
void ChangeFontSize(float delta);
void ChangeAlignment(TextAlignment align);
void BringForward();
void SendBackward();
void DrawShapePath(GraphicsPath& path, ShapeType type, int x, int y, int w, int h);
void DrawElement(Graphics& g, const SlideElement& elem, int slideX, int slideY, int slideW, int slideH, bool isThumbnail);

// ==================== INITIALIZATION ====================
void InitializeRibbon() {
    g_state.ribbonTabs.clear();
    
    // Home Tab
    RibbonTab homeTab;
    homeTab.name = L"Home";
    homeTab.isActive = true;
    
    // Clipboard Group
    RibbonGroup clipboardGroup;
    clipboardGroup.name = L"Clipboard";
    clipboardGroup.buttons.push_back({L"Paste", L"📋", {0, 0, 0, 0}, false, false, 100});
    clipboardGroup.buttons.push_back({L"Cut", L"✂️", {0, 0, 0, 0}, false, false, 101});
    clipboardGroup.buttons.push_back({L"Copy", L"📄", {0, 0, 0, 0}, false, false, 102});
    homeTab.groups.push_back(clipboardGroup);
    
    // Slides Group
    RibbonGroup slidesGroup;
    slidesGroup.name = L"Slides";
    slidesGroup.buttons.push_back({L"New Slide", L"➕", {0, 0, 0, 0}, false, false, 200});
    slidesGroup.buttons.push_back({L"Delete", L"🗑️", {0, 0, 0, 0}, false, false, 201});
    slidesGroup.buttons.push_back({L"Duplicate", L"📑", {0, 0, 0, 0}, false, false, 202});
    homeTab.groups.push_back(slidesGroup);
    
    // Font Group
    RibbonGroup fontGroup;
    fontGroup.name = L"Font";
    fontGroup.buttons.push_back({L"Bold", L"𝐁", {0, 0, 0, 0}, false, false, 300});
    fontGroup.buttons.push_back({L"Italic", L"𝐼", {0, 0, 0, 0}, false, false, 301});
    fontGroup.buttons.push_back({L"Underline", L"U̲", {0, 0, 0, 0}, false, false, 302});
    fontGroup.buttons.push_back({L"Bigger", L"A▲", {0, 0, 0, 0}, false, false, 303});
    fontGroup.buttons.push_back({L"Smaller", L"A▼", {0, 0, 0, 0}, false, false, 304});
    homeTab.groups.push_back(fontGroup);
    
    // Paragraph Group
    RibbonGroup paragraphGroup;
    paragraphGroup.name = L"Paragraph";
    paragraphGroup.buttons.push_back({L"Left", L"⫷", {0, 0, 0, 0}, false, false, 400});
    paragraphGroup.buttons.push_back({L"Center", L"≣", {0, 0, 0, 0}, false, false, 401});
    paragraphGroup.buttons.push_back({L"Right", L"⫸", {0, 0, 0, 0}, false, false, 402});
    homeTab.groups.push_back(paragraphGroup);
    
    // Drawing Group
    RibbonGroup drawingGroup;
    drawingGroup.name = L"Drawing";
    drawingGroup.buttons.push_back({L"Rectangle", L"▬", {0, 0, 0, 0}, false, false, 500});
    drawingGroup.buttons.push_back({L"Rounded", L"▭", {0, 0, 0, 0}, false, false, 501});
    drawingGroup.buttons.push_back({L"Ellipse", L"⬭", {0, 0, 0, 0}, false, false, 502});
    drawingGroup.buttons.push_back({L"Triangle", L"▲", {0, 0, 0, 0}, false, false, 503});
    drawingGroup.buttons.push_back({L"Star", L"★", {0, 0, 0, 0}, false, false, 504});
    homeTab.groups.push_back(drawingGroup);
    
    // Arrange Group
    RibbonGroup arrangeGroup;
    arrangeGroup.name = L"Arrange";
    arrangeGroup.buttons.push_back({L"Forward", L"⬆", {0, 0, 0, 0}, false, false, 600});
    arrangeGroup.buttons.push_back({L"Backward", L"⬇", {0, 0, 0, 0}, false, false, 601});
    homeTab.groups.push_back(arrangeGroup);
    
    g_state.ribbonTabs.push_back(homeTab);
    
    // Insert Tab
    RibbonTab insertTab;
    insertTab.name = L"Insert";
    insertTab.isActive = false;
    
    RibbonGroup insertGroup;
    insertGroup.name = L"Insert";
    insertGroup.buttons.push_back({L"Text Box", L"📝", {0, 0, 0, 0}, false, false, 700});
    insertGroup.buttons.push_back({L"Picture", L"🖼️", {0, 0, 0, 0}, false, false, 701});
    insertGroup.buttons.push_back({L"Chart", L"📊", {0, 0, 0, 0}, false, false, 702});
    insertTab.groups.push_back(insertGroup);
    
    g_state.ribbonTabs.push_back(insertTab);
    
    // Design Tab
    RibbonTab designTab;
    designTab.name = L"Design";
    designTab.isActive = false;
    
    RibbonGroup themesGroup;
    themesGroup.name = L"Themes";
    themesGroup.buttons.push_back({L"White", L"⬜", {0, 0, 0, 0}, false, false, 800});
    themesGroup.buttons.push_back({L"Blue", L"🟦", {0, 0, 0, 0}, false, false, 801});
    themesGroup.buttons.push_back({L"Dark", L"⬛", {0, 0, 0, 0}, false, false, 802});
    designTab.groups.push_back(themesGroup);
    
    g_state.ribbonTabs.push_back(designTab);
}

void InitializeDefaultSlides() {
    g_state.slides.clear();
    
    // Slide 1 - Title Slide
    Slide slide1;
    slide1.id = 1;
    slide1.bgColor = RGB(255, 255, 255);
    slide1.title = L"Title Slide";
    
    SlideElement title1;
    title1.id = 1;
    title1.rect = {15, 20, 85, 45};
    title1.text = L"Welcome to PowerPoint Editor";
    title1.fontSize = 44.0f;
    title1.isBold = true;
    title1.textColor = RGB(44, 62, 80);
    title1.bgColor = RGB(236, 240, 241);
    title1.shapeType = SHAPE_ROUNDED_RECT;
    title1.borderColor = RGB(52, 152, 219);
    title1.borderWidth = 3;
    title1.alignment = ALIGN_CENTER;
    slide1.elements.push_back(title1);
    
    SlideElement subtitle1;
    subtitle1.id = 2;
    subtitle1.rect = {20, 50, 80, 65};
    subtitle1.text = L"Professional Presentation Editor with Ribbon Interface";
    subtitle1.fontSize = 24.0f;
    subtitle1.textColor = RGB(127, 140, 141);
    subtitle1.alignment = ALIGN_CENTER;
    slide1.elements.push_back(subtitle1);
    
    g_state.slides.push_back(slide1);
    
    // Slide 2 - Content Slide
    Slide slide2;
    slide2.id = 2;
    slide2.bgColor = RGB(248, 249, 250);
    slide2.title = L"Features";
    
    SlideElement header2;
    header2.id = 3;
    header2.rect = {10, 10, 90, 30};
    header2.text = L"Key Features";
    header2.fontSize = 36.0f;
    header2.isBold = true;
    header2.textColor = RGB(41, 128, 185);
    header2.shapeType = SHAPE_RECTANGLE;
    header2.bgColor = RGB(214, 234, 248);
    header2.borderColor = RGB(41, 128, 185);
    header2.borderWidth = 2;
    header2.alignment = ALIGN_CENTER;
    slide2.elements.push_back(header2);
    
    SlideElement bullet1;
    bullet1.id = 4;
    bullet1.rect = {15, 38, 85, 50};
    bullet1.text = L"✓ Ribbon-based interface for easy access to tools";
    bullet1.fontSize = 20.0f;
    bullet1.textColor = RGB(44, 62, 80);
    slide2.elements.push_back(bullet1);
    
    SlideElement bullet2;
    bullet2.id = 5;
    bullet2.rect = {15, 52, 85, 64};
    bullet2.text = L"✓ Multiple shapes with customizable colors and borders";
    bullet2.fontSize = 20.0f;
    bullet2.textColor = RGB(44, 62, 80);
    slide2.elements.push_back(bullet2);
    
    SlideElement bullet3;
    bullet3.id = 6;
    bullet3.rect = {15, 66, 85, 78};
    bullet3.text = L"✓ Full text formatting (Bold, Italic, Underline, Alignment)";
    bullet3.fontSize = 20.0f;
    bullet3.textColor = RGB(44, 62, 80);
    slide2.elements.push_back(bullet3);
    
    SlideElement bullet4;
    bullet4.id = 7;
    bullet4.rect = {15, 80, 85, 92};
    bullet4.text = L"✓ Undo/Redo support and slide management";
    bullet4.fontSize = 20.0f;
    bullet4.textColor = RGB(44, 62, 80);
    slide2.elements.push_back(bullet4);
    
    g_state.slides.push_back(slide2);
    
    // Slide 3 - Graphics Demo
    Slide slide3;
    slide3.id = 3;
    slide3.bgColor = RGB(255, 255, 255);
    slide3.title = L"Graphics";
    
    SlideElement shape1;
    shape1.id = 8;
    shape1.rect = {10, 30, 35, 55};
    shape1.text = L"Rectangle";
    shape1.fontSize = 16.0f;
    shape1.textColor = RGB(255, 255, 255);
    shape1.bgColor = RGB(231, 76, 60);
    shape1.shapeType = SHAPE_RECTANGLE;
    shape1.alignment = ALIGN_CENTER;
    slide3.elements.push_back(shape1);
    
    SlideElement shape2;
    shape2.id = 9;
    shape2.rect = {40, 30, 65, 55};
    shape2.text = L"Ellipse";
    shape2.fontSize = 16.0f;
    shape2.textColor = RGB(255, 255, 255);
    shape2.bgColor = RGB(46, 204, 113);
    shape2.shapeType = SHAPE_ELLIPSE;
    shape2.alignment = ALIGN_CENTER;
    slide3.elements.push_back(shape2);
    
    SlideElement shape3;
    shape3.id = 10;
    shape3.rect = {70, 30, 95, 55};
    shape3.text = L"Star";
    shape3.fontSize = 16.0f;
    shape3.textColor = RGB(255, 255, 255);
    shape3.bgColor = RGB(241, 196, 15);
    shape3.shapeType = SHAPE_STAR;
    shape3.alignment = ALIGN_CENTER;
    slide3.elements.push_back(shape3);
    
    SlideElement shape4;
    shape4.id = 11;
    shape4.rect = {25, 65, 50, 85};
    shape4.text = L"Diamond";
    shape4.fontSize = 16.0f;
    shape4.textColor = RGB(255, 255, 255);
    shape4.bgColor = RGB(155, 89, 182);
    shape4.shapeType = SHAPE_DIAMOND;
    shape4.alignment = ALIGN_CENTER;
    slide3.elements.push_back(shape4);
    
    SlideElement shape5;
    shape5.id = 12;
    shape5.rect = {55, 65, 80, 85};
    shape5.text = L"Triangle";
    shape5.fontSize = 16.0f;
    shape5.textColor = RGB(255, 255, 255);
    shape5.bgColor = RGB(52, 152, 219);
    shape5.shapeType = SHAPE_TRIANGLE;
    shape5.alignment = ALIGN_CENTER;
    slide3.elements.push_back(shape5);
    
    g_state.slides.push_back(slide3);
    
    g_state.currentSlide = 0;
}

// ==================== STATE MANAGEMENT ====================
void SaveState() {
    // Remove any redo states
    while ((int)g_state.undoStack.size() > g_state.undoIndex + 1) {
        g_state.undoStack.pop_back();
    }
    
    g_state.undoStack.push_back(g_state.slides);
    g_state.undoIndex = (int)g_state.undoStack.size() - 1;
    g_state.isModified = true;
}

void Undo() {
    if (g_state.undoIndex > 0) {
        g_state.undoIndex--;
        g_state.slides = g_state.undoStack[g_state.undoIndex];
        g_state.currentSlide = min(g_state.currentSlide, (int)g_state.slides.size() - 1);
        g_state.selectedElementId = -1;
        InvalidateRect(g_hwnd, NULL, TRUE);
    }
}

void Redo() {
    if (g_state.undoIndex < (int)g_state.undoStack.size() - 1) {
        g_state.undoIndex++;
        g_state.slides = g_state.undoStack[g_state.undoIndex];
        g_state.currentSlide = min(g_state.currentSlide, (int)g_state.slides.size() - 1);
        g_state.selectedElementId = -1;
        InvalidateRect(g_hwnd, NULL, TRUE);
    }
}

// ==================== SLIDE OPERATIONS ====================
void AddNewSlide() {
    SaveState();
    
    Slide newSlide;
    newSlide.id = g_state.slides.size() + 1;
    newSlide.title = L"Slide " + to_wstring(newSlide.id);
    newSlide.bgColor = RGB(255, 255, 255);
    
    g_state.slides.insert(g_state.slides.begin() + g_state.currentSlide + 1, newSlide);
    g_state.currentSlide++;
    g_state.selectedElementId = -1;
    InvalidateRect(g_hwnd, NULL, TRUE);
}

void DeleteCurrentSlide() {
    if (g_state.slides.size() > 1) {
        SaveState();
        g_state.slides.erase(g_state.slides.begin() + g_state.currentSlide);
        g_state.currentSlide = min(g_state.currentSlide, (int)g_state.slides.size() - 1);
        g_state.selectedElementId = -1;
        InvalidateRect(g_hwnd, NULL, TRUE);
    }
}

void AddTextBox() {
    if (!g_state.slides.empty()) {
        SaveState();
        
        SlideElement textBox;
        textBox.id = rand();
        textBox.rect = {25, 25, 75, 45};
        textBox.text = L"New Text Box";
        textBox.fontSize = 24.0f;
        textBox.textColor = RGB(44, 62, 80);
        textBox.bgColor = RGB(236, 240, 241);
        textBox.borderColor = RGB(189, 195, 199);
        textBox.borderWidth = 1;
        textBox.shapeType = SHAPE_RECTANGLE;
        
        g_state.slides[g_state.currentSlide].elements.push_back(textBox);
        g_state.selectedElementId = textBox.id;
        InvalidateRect(g_hwnd, NULL, TRUE);
    }
}

void AddShape(ShapeType type) {
    if (!g_state.slides.empty()) {
        SaveState();
        
        SlideElement shape;
        shape.id = rand();
        shape.rect = {30, 30, 70, 70};
        shape.text = L"";
        shape.fontSize = 18.0f;
        shape.textColor = RGB(255, 255, 255);
        shape.bgColor = RGB(52, 152, 219);
        shape.shapeType = type;
        
        g_state.slides[g_state.currentSlide].elements.push_back(shape);
        g_state.selectedElementId = shape.id;
        InvalidateRect(g_hwnd, NULL, TRUE);
    }
}

void ChangeFillColor(COLORREF color) {
    if (g_state.currentSlide < (int)g_state.slides.size() && g_state.selectedElementId >= 0) {
        SaveState();
        for (auto& elem : g_state.slides[g_state.currentSlide].elements) {
            if (elem.id == g_state.selectedElementId) {
                elem.bgColor = color;
                break;
            }
        }
        InvalidateRect(g_hwnd, NULL, TRUE);
    }
}

void ChangeTextColor(COLORREF color) {
    if (g_state.currentSlide < (int)g_state.slides.size() && g_state.selectedElementId >= 0) {
        SaveState();
        for (auto& elem : g_state.slides[g_state.currentSlide].elements) {
            if (elem.id == g_state.selectedElementId) {
                elem.textColor = color;
                break;
            }
        }
        InvalidateRect(g_hwnd, NULL, TRUE);
    }
}


void ChangeBorderColor(COLORREF color) {
    if (g_state.currentSlide < (int)g_state.slides.size() && g_state.selectedElementId >= 0) {
        SaveState();
        for (auto& elem : g_state.slides[g_state.currentSlide].elements) {
            if (elem.id == g_state.selectedElementId) {
                elem.borderColor = color;
                break;
            }
        }
        InvalidateRect(g_hwnd, NULL, TRUE);
    }
}

void ToggleBold() {
    if (g_state.currentSlide < (int)g_state.slides.size() && g_state.selectedElementId >= 0) {
        SaveState();
        for (auto& elem : g_state.slides[g_state.currentSlide].elements) {
            if (elem.id == g_state.selectedElementId) {
                elem.isBold = !elem.isBold;
                break;
            }
        }
        InvalidateRect(g_hwnd, NULL, TRUE);
    }
}

void ToggleItalic() {
    if (g_state.currentSlide < (int)g_state.slides.size() && g_state.selectedElementId >= 0) {
        SaveState();
        for (auto& elem : g_state.slides[g_state.currentSlide].elements) {
            if (elem.id == g_state.selectedElementId) {
                elem.isItalic = !elem.isItalic;
                break;
            }
        }
        InvalidateRect(g_hwnd, NULL, TRUE);
    }
}


void ToggleUnderline() {
    if (g_state.currentSlide < (int)g_state.slides.size() && g_state.selectedElementId >= 0) {
        SaveState();
        for (auto& elem : g_state.slides[g_state.currentSlide].elements) {
            if (elem.id == g_state.selectedElementId) {
                elem.isUnderline = !elem.isUnderline;
                break;
            }
        }
        InvalidateRect(g_hwnd, NULL, TRUE);
    }
}

void ChangeFontSize(float delta) {
    if (g_state.currentSlide < (int)g_state.slides.size() && g_state.selectedElementId >= 0) {
        SaveState();
        for (auto& elem : g_state.slides[g_state.currentSlide].elements) {
            if (elem.id == g_state.selectedElementId) {
                elem.fontSize = max(8.0f, min(96.0f, elem.fontSize + delta));
                break;
            }
        }
        InvalidateRect(g_hwnd, NULL, TRUE);
    }
}

void ChangeAlignment(TextAlignment align) {
    if (g_state.currentSlide < (int)g_state.slides.size() && g_state.selectedElementId >= 0) {
        SaveState();
        for (auto& elem : g_state.slides[g_state.currentSlide].elements) {
            if (elem.id == g_state.selectedElementId) {
                elem.alignment = align;
                break;
            }
        }
        InvalidateRect(g_hwnd, NULL, TRUE);
    }
}

void BringForward() {
    if (g_state.currentSlide < (int)g_state.slides.size() && g_state.selectedElementId >= 0) {
        SaveState();
        auto& elements = g_state.slides[g_state.currentSlide].elements;
        for (size_t i = 0; i < elements.size() - 1; i++) {
            if (elements[i].id == g_state.selectedElementId) {
                swap(elements[i], elements[i + 1]);
                break;
            }
        }
        InvalidateRect(g_hwnd, NULL, TRUE);
    }
}

void SendBackward() {
    if (g_state.currentSlide < (int)g_state.slides.size() && g_state.selectedElementId >= 0) {
        SaveState();
        auto& elements = g_state.slides[g_state.currentSlide].elements;
        for (size_t i = 1; i < elements.size(); i++) {
            if (elements[i].id == g_state.selectedElementId) {
                swap(elements[i], elements[i - 1]);
                break;
            }
        }
        InvalidateRect(g_hwnd, NULL, TRUE);
    }
}

// ==================== DRAWING FUNCTIONS ====================
void DrawShapePath(GraphicsPath& path, ShapeType type, int x, int y, int w, int h) {
    switch (type) {
        case SHAPE_RECTANGLE:
            path.AddRectangle(Rect(x, y, w, h));
            break;
        case SHAPE_ROUNDED_RECT: {
            int r = 15;
            path.AddArc(x, y, r, r, 180, 90);
            path.AddArc(x + w - r, y, r, r, 270, 90);
            path.AddArc(x + w - r, y + h - r, r, r, 0, 90);
            path.AddArc(x, y + h - r, r, r, 90, 90);
            path.CloseFigure();
            break;
        }
        case SHAPE_ELLIPSE:
            path.AddEllipse(x, y, w, h);
            break;
        case SHAPE_TRIANGLE: {
            Point pts[3] = {
                {x + w/2, y},
                {x + w, y + h},
                {x, y + h}
            };
            path.AddPolygon(pts, 3);
            break;
        }
        case SHAPE_DIAMOND: {
            Point pts[4] = {
                {x + w/2, y},
                {x + w, y + h/2},
                {x + w/2, y + h},
                {x, y + h/2}
            };
            path.AddPolygon(pts, 4);
            break;
        }
        case SHAPE_ARROW_RIGHT:
            path.AddRectangle(Rect(x, y, w, h)); // fallback
            break;
        case SHAPE_STAR: {
            int cx = x + w/2, cy = y + h/2;
            int outerR = min(w, h) / 2;
            int innerR = outerR / 2;
            Point pts[10];
            for (int i = 0; i < 10; i++) {
                double angle = -M_PI/2 + i * M_PI/5;
                int r = (i % 2 == 0) ? outerR : innerR;
                pts[i] = {cx + (int)(r * cos(angle)), cy + (int)(r * sin(angle))};
            }
            path.AddPolygon(pts, 10);
            break;
        }
    }
}

void DrawRibbon(Graphics& g, int width) {
    int ribbonH = g_state.ribbonHeight;
    
    // Ribbon background
    SolidBrush ribbonBg(Color(255, 243, 246, 250));
    g.FillRectangle(&ribbonBg, 0, 0, width, ribbonH);
    
    // Bottom border
    Pen borderPen(Color(255, 200, 200, 210), 1.0f);
    g.DrawLine(&borderPen, 0, ribbonH - 1, width, ribbonH - 1);
    
    // Draw tabs
    int tabY = 0;
    int tabH = 25;
    int tabX = 10;
    
    Font tabFont(L"Segoe UI", 11.0f * g_scaleFactor, FontStyleRegular, UnitPoint);
    StringFormat tabFormat;
    tabFormat.SetAlignment(StringAlignmentCenter);
    tabFormat.SetLineAlignment(StringAlignmentCenter);
    
    for (size_t i = 0; i < g_state.ribbonTabs.size(); i++) {
        int tabW = 80;
        g_state.ribbonTabs[i].tabRect = {tabX, tabY, tabX + tabW, tabY + tabH};
        
        if (g_state.ribbonTabs[i].isActive) {
            SolidBrush activeTab(Color(255, 255, 255, 255));
            g.FillRectangle(&activeTab, tabX, tabY, tabW, tabH);
            Pen activeBorder(Color(255, 41, 128, 185), 2.0f);
            g.DrawLine(&activeBorder, tabX, tabY + tabH - 2, tabX + tabW, tabY + tabH - 2);
        }
        
        SolidBrush tabTextBrush(Color(255, 44, 62, 80));
        RectF tabRectF((float)tabX, (float)tabY, (float)tabW, (float)tabH);
        g.DrawString(g_state.ribbonTabs[i].name.c_str(), -1, &tabFont, tabRectF, &tabFormat, &tabTextBrush);
        
        tabX += tabW + 5;
    }
    
    // Draw groups for active tab
    if (g_state.activeTabIndex < (int)g_state.ribbonTabs.size()) {
        auto& activeTab = g_state.ribbonTabs[g_state.activeTabIndex];
        int groupX = 10;
        int groupY = tabH + 10;
        int buttonW = 55;
        int buttonH = 45;
        
        Font btnFont(L"Segoe UI", 9.0f * g_scaleFactor, FontStyleRegular, UnitPoint);
        Font iconFont(L"Segoe UI Symbol", 14.0f * g_scaleFactor, FontStyleRegular, UnitPoint);
        StringFormat btnFormat;
        btnFormat.SetAlignment(StringAlignmentCenter);
        btnFormat.SetLineAlignment(StringAlignmentCenter);
        
        for (auto& group : activeTab.groups) {
            int groupW = (int)(group.buttons.size() * (buttonW + 5) + 15);
            group.rect = {groupX, groupY, groupX + groupW, ribbonH - 5};
            
            // Group separator
            Pen sepPen(Color(255, 220, 220, 225), 1.0f);
            g.DrawLine(&sepPen, groupX, groupY, groupX, ribbonH - 5);
            
            // Group name
            Font groupFont(L"Segoe UI", 8.0f * g_scaleFactor, FontStyleRegular, UnitPoint);
            SolidBrush groupBrush(Color(255, 150, 150, 160));
            RectF groupRectF((float)(groupX + 5), (float)(ribbonH - 15), (float)(groupW - 10), 12.0f);
            g.DrawString(group.name.c_str(), -1, &groupFont, groupRectF, &btnFormat, &groupBrush);
            
            int btnX = groupX + 5;
            int btnY = groupY + 5;
            
            for (auto& button : group.buttons) {
                button.rect = {btnX, btnY, btnX + buttonW, btnY + buttonH};
                
                // Button background
                if (button.isHovered) {
                    SolidBrush hoverBrush(Color(255, 214, 234, 248));
                    g.FillRectangle(&hoverBrush, btnX, btnY, buttonW, buttonH);
                    Pen hoverBorder(Color(255, 41, 128, 185), 1.0f);
                    g.DrawRectangle(&hoverBorder, btnX, btnY, buttonW, buttonH);
                }
                
                // Icon
                RectF iconRect((float)btnX, (float)btnY, (float)buttonW, 30.0f);
                SolidBrush btnTextBrush(Color(255, 44, 62, 80));
                g.DrawString(button.icon.c_str(), -1, &iconFont, iconRect, &btnFormat, &btnTextBrush);
                
                // Text
                RectF textRect((float)btnX, (float)(btnY + 28), (float)buttonW, (float)(buttonH - 28));
                g.DrawString(button.text.c_str(), -1, &btnFont, textRect, &btnFormat, &btnTextBrush);
                
                btnX += buttonW + 3;
            }
            
            groupX += groupW + 10;
        }
    }
}

void DrawSidebar(Graphics& g, int width, int height) {
    int sidebarW = g_state.sidebarWidth;
    int ribbonH = g_state.ribbonHeight;
    
    // Sidebar background
    SolidBrush sidebarBg(Color(255, 236, 240, 241));
    g.FillRectangle(&sidebarBg, 0, ribbonH, sidebarW, height - ribbonH);
    
    // Border
    Pen borderPen(Color(255, 189, 195, 199), 1.0f);
    g.DrawLine(&borderPen, sidebarW - 1, ribbonH, sidebarW - 1, height);
    
    // "Slides" label
    Font labelFont(L"Segoe UI", 10.0f, FontStyleBold, UnitPoint);
    SolidBrush labelBrush(Color(255, 80, 80, 100));
    g.DrawString(L"SLIDES", -1, &labelFont, RectF(8, (float)(ribbonH + 6), (float)(sidebarW - 16), 18), nullptr, &labelBrush);
    
    int thumbW = sidebarW - 20;
    int thumbH = (int)(thumbW * 9.0f / 16.0f);
    int thumbX = 10;
    int thumbY = ribbonH + 28 - g_state.thumbnailScrollPos;
    
    for (int i = 0; i < (int)g_state.slides.size(); i++) {
        if (thumbY + thumbH >= ribbonH && thumbY < height) {
            // Selected highlight
            if (i == g_state.currentSlide) {
                SolidBrush selBrush(Color(255, 214, 234, 248));
                g.FillRectangle(&selBrush, thumbX - 4, thumbY - 4, thumbW + 8, thumbH + 8);
                Pen selPen(Color(255, 41, 128, 185), 2.0f);
                g.DrawRectangle(&selPen, thumbX - 4, thumbY - 4, thumbW + 8, thumbH + 8);
            }
            
            // Draw thumbnail
            DrawSlide(g, thumbX, thumbY, thumbW, thumbH, g_state.slides[i], true);
            
            // Slide number
            Font numFont(L"Segoe UI", 8.0f, FontStyleRegular, UnitPoint);
            SolidBrush numBrush(Color(255, 120, 120, 140));
            wstring numStr = to_wstring(i + 1);
            g.DrawString(numStr.c_str(), -1, &numFont,
                PointF((float)(thumbX - 2), (float)(thumbY + thumbH / 2 - 6)), &numBrush);
        }
        thumbY += thumbH + 12;
    }
}

void DrawElement(Graphics& g, const SlideElement& elem, int slideX, int slideY, int slideW, int slideH, bool isThumbnail) {
    // Convert % coords to pixels
    int x = slideX + (int)(elem.rect.left   * slideW / 100.0f);
    int y = slideY + (int)(elem.rect.top    * slideH / 100.0f);
    int w =          (int)((elem.rect.right  - elem.rect.left) * slideW / 100.0f);
    int h =          (int)((elem.rect.bottom - elem.rect.top)  * slideH / 100.0f);
    if (w <= 0) w = 1;
    if (h <= 0) h = 1;

    BYTE bgR = GetRValue(elem.bgColor), bgG = GetGValue(elem.bgColor), bgB = GetBValue(elem.bgColor);
    BYTE txR = GetRValue(elem.textColor), txG = GetGValue(elem.textColor), txB = GetBValue(elem.textColor);
    BYTE bdR = GetRValue(elem.borderColor), bdG = GetGValue(elem.borderColor), bdB = GetBValue(elem.borderColor);

    SolidBrush fillBrush(Color(255, bgR, bgG, bgB));
    Pen        borderPen(Color(255, bdR, bdG, bdB), (float)elem.borderWidth);
    SolidBrush textBrush(Color(255, txR, txG, txB));

    // Build shape path
    GraphicsPath path;
    DrawShapePath(path, elem.shapeType, x, y, w, h);
    g.FillPath(&fillBrush, &path);
    if (elem.borderWidth > 0)
        g.DrawPath(&borderPen, &path);

    // Text
    if (!elem.text.empty()) {
        float scaledSize = isThumbnail ? max(5.0f, elem.fontSize * slideW / 960.0f)
                                       : elem.fontSize * g_scaleFactor;
        int style = FontStyleRegular;
        if (elem.isBold)      style |= FontStyleBold;
        if (elem.isItalic)    style |= FontStyleItalic;
        if (elem.isUnderline) style |= FontStyleUnderline;

        Font textFont(L"Segoe UI", scaledSize, style, UnitPoint);
        StringFormat fmt;
        switch (elem.alignment) {
            case ALIGN_LEFT:   fmt.SetAlignment(StringAlignmentNear);  break;
            case ALIGN_CENTER: fmt.SetAlignment(StringAlignmentCenter); break;
            case ALIGN_RIGHT:  fmt.SetAlignment(StringAlignmentFar);    break;
        }
        fmt.SetLineAlignment(StringAlignmentCenter);
        fmt.SetTrimming(StringTrimmingEllipsisWord);

        RectF textRect((float)(x + 6), (float)(y + 4), (float)(w - 12), (float)(h - 8));
        g.DrawString(elem.text.c_str(), -1, &textFont, textRect, &fmt, &textBrush);
    }

    // Selection handles
    if (elem.isSelected && !isThumbnail) {
        Pen selPen(Color(255, 41, 128, 185), 1.5f);
        float dash[] = {4.0f, 2.0f};
        selPen.SetDashPattern(dash, 2);
        g.DrawRectangle(&selPen, x - 2, y - 2, w + 4, h + 4);

        SolidBrush handleBrush(Color(255, 41, 128, 185));
        int hs = 7;
        int hx[] = {x-hs/2, x+w/2-hs/2, x+w-hs/2, x-hs/2, x+w-hs/2, x-hs/2, x+w/2-hs/2, x+w-hs/2};
        int hy[] = {y-hs/2, y-hs/2, y-hs/2, y+h/2-hs/2, y+h/2-hs/2, y+h-hs/2, y+h-hs/2, y+h-hs/2};
        for (int k = 0; k < 8; k++)
            g.FillRectangle(&handleBrush, hx[k], hy[k], hs, hs);
    }
}

void DrawSlide(Graphics& g, int x, int y, int width, int height, const Slide& slide, bool isThumbnail) {
    // Background
    BYTE r = GetRValue(slide.bgColor), gr = GetGValue(slide.bgColor), b = GetBValue(slide.bgColor);
    SolidBrush bgBrush(Color(255, r, gr, b));
    g.FillRectangle(&bgBrush, x, y, width, height);

    // Border
    Pen borderPen(Color(255, 180, 180, 190), 1.0f);
    g.DrawRectangle(&borderPen, x, y, width, height);

    // Shadow (main slide only)
    if (!isThumbnail) {
        SolidBrush shadowBrush(Color(40, 0, 0, 0));
        g.FillRectangle(&shadowBrush, x + 4, y + 4, width, height);
    }

    // Clip to slide area
    g.SetClip(Rect(x, y, width, height));
    for (const auto& elem : slide.elements)
        DrawElement(g, elem, x, y, width, height, isThumbnail);
    g.ResetClip();
}

void DrawStatusBar(Graphics& g, int width, int height) {
    int sbY = height - g_state.statusBarHeight;
    SolidBrush sbBrush(Color(255, 41, 128, 185));
    g.FillRectangle(&sbBrush, 0, sbY, width, g_state.statusBarHeight);

    Font sbFont(L"Segoe UI", 9.0f, FontStyleRegular, UnitPoint);
    SolidBrush sbText(Color(255, 255, 255, 255));

    wstring status = L"  Slide " + to_wstring(g_state.currentSlide + 1) +
                     L" of " + to_wstring(g_state.slides.size());
    if (g_state.selectedElementId >= 0)
        status += L"  |  Element selected";
    if (g_state.isModified)
        status += L"  |  Modified";
    if (!g_state.currentFile.empty())
        status += L"  |  " + g_state.currentFile;

    g.DrawString(status.c_str(), -1, &sbFont, PointF(4.0f, (float)(sbY + 4)), &sbText);

    // Right side: zoom
    wstring zoomStr = to_wstring((int)(g_scaleFactor * 100)) + L"%";
    StringFormat rightFmt;
    rightFmt.SetAlignment(StringAlignmentFar);
    g.DrawString(zoomStr.c_str(), -1, &sbFont,
        RectF(0, (float)(sbY + 4), (float)(width - 8), 16), &rightFmt, &sbText);
}

void DrawColorPicker(Graphics& g, int x, int y) {
    int cols = 8, rows = 5, cellSize = 24, padding = 4;
    int panelW = cols * (cellSize + padding) + padding + 4;
    int panelH = rows * (cellSize + padding) + padding + 4;

    SolidBrush panelBg(Color(255, 255, 255, 255));
    g.FillRectangle(&panelBg, x, y, panelW, panelH);
    Pen panelBorder(Color(255, 180, 180, 190), 1.0f);
    g.DrawRectangle(&panelBorder, x, y, panelW, panelH);

    COLORREF colors[40] = {
        RGB(0,0,0),       RGB(64,64,64),    RGB(128,128,128), RGB(192,192,192),
        RGB(255,255,255), RGB(231,76,60),   RGB(192,57,43),   RGB(241,196,15),
        RGB(243,156,18),  RGB(46,204,113),  RGB(39,174,96),   RGB(26,188,156),
        RGB(22,160,133),  RGB(52,152,219),  RGB(41,128,185),  RGB(155,89,182),
        RGB(142,68,173),  RGB(52,73,94),    RGB(44,62,80),    RGB(236,240,241),
        RGB(255,100,100), RGB(100,200,100), RGB(100,100,255), RGB(255,200,0),
        RGB(200,100,200), RGB(0,200,200),   RGB(255,150,50),  RGB(150,50,200),
        RGB(50,150,100),  RGB(200,50,50),   RGB(50,200,150),  RGB(100,150,200),
        RGB(200,150,100), RGB(150,200,50),  RGB(255,128,0),   RGB(128,0,255),
        RGB(0,128,255),   RGB(255,0,128),   RGB(0,255,128),   RGB(128,255,0)
    };

    for (int i = 0; i < 40; i++) {
        int col = i % cols, row = i / cols;
        int cx = x + padding + col * (cellSize + padding);
        int cy = y + padding + row * (cellSize + padding);

        BYTE cr = GetRValue(colors[i]), cg = GetGValue(colors[i]), cb = GetBValue(colors[i]);
        SolidBrush colorBrush(Color(255, cr, cg, cb));
        g.FillRectangle(&colorBrush, cx, cy, cellSize, cellSize);
        Pen cellBorder(Color(255, 180, 180, 190), 1.0f);
        g.DrawRectangle(&cellBorder, cx, cy, cellSize, cellSize);
    }
}

// ==================== HIT TEST ====================
void GetSlideArea(int winW, int winH, int& slideX, int& slideY, int& slideW, int& slideH) {
    int contentX = g_state.sidebarWidth;
    int contentY = g_state.ribbonHeight;
    int contentW = winW - g_state.sidebarWidth;
    int contentH = winH - g_state.ribbonHeight - g_state.statusBarHeight;

    float slideAspect = 16.0f / 9.0f;
    int targetW = contentW - 60;
    int targetH = (int)(targetW / slideAspect);
    if (targetH > contentH - 60) {
        targetH = contentH - 60;
        targetW = (int)(targetH * slideAspect);
    }
    slideX = contentX + (contentW - targetW) / 2;
    slideY = contentY + (contentH - targetH) / 2;
    slideW = targetW;
    slideH = targetH;
}

// ==================== EVENT HANDLERS ====================
void HandleRibbonClick(int x, int y) {
    // Check tab clicks
    for (size_t i = 0; i < g_state.ribbonTabs.size(); i++) {
        RECT& tr = g_state.ribbonTabs[i].tabRect;
        if (x >= tr.left && x < tr.right && y >= tr.top && y < tr.bottom) {
            for (auto& tab : g_state.ribbonTabs) tab.isActive = false;
            g_state.ribbonTabs[i].isActive = true;
            g_state.activeTabIndex = (int)i;
            InvalidateRect(g_hwnd, NULL, TRUE);
            return;
        }
    }

    if (g_state.activeTabIndex >= (int)g_state.ribbonTabs.size()) return;
    auto& activeTab = g_state.ribbonTabs[g_state.activeTabIndex];

    for (auto& group : activeTab.groups) {
        for (auto& btn : group.buttons) {
            RECT& br = btn.rect;
            if (x >= br.left && x < br.right && y >= br.top && y < br.bottom) {
                switch (btn.id) {
                    case 100: /* Paste */  break;
                    case 101: /* Cut */    break;
                    case 102: /* Copy */   break;
                    case 200: AddNewSlide(); break;
                    case 201: DeleteCurrentSlide(); break;
                    case 202: {
                        if (!g_state.slides.empty()) {
                            SaveState();
                            Slide dup = g_state.slides[g_state.currentSlide];
                            dup.id = (int)g_state.slides.size() + 1;
                            g_state.slides.insert(g_state.slides.begin() + g_state.currentSlide + 1, dup);
                            g_state.currentSlide++;
                            InvalidateRect(g_hwnd, NULL, TRUE);
                        }
                        break;
                    }
                    case 300: ToggleBold(); break;
                    case 301: ToggleItalic(); break;
                    case 302: ToggleUnderline(); break;
                    case 303: ChangeFontSize(2.0f); break;
                    case 304: ChangeFontSize(-2.0f); break;
                    case 400: ChangeAlignment(ALIGN_LEFT); break;
                    case 401: ChangeAlignment(ALIGN_CENTER); break;
                    case 402: ChangeAlignment(ALIGN_RIGHT); break;
                    case 500: AddShape(SHAPE_RECTANGLE); break;
                    case 501: AddShape(SHAPE_ROUNDED_RECT); break;
                    case 502: AddShape(SHAPE_ELLIPSE); break;
                    case 503: AddShape(SHAPE_TRIANGLE); break;
                    case 504: AddShape(SHAPE_STAR); break;
                    case 600: BringForward(); break;
                    case 601: SendBackward(); break;
                    case 700: AddTextBox(); break;
                    case 800: {
                        if (!g_state.slides.empty()) {
                            g_state.slides[g_state.currentSlide].bgColor = RGB(255,255,255);
                            InvalidateRect(g_hwnd, NULL, TRUE);
                        }
                        break;
                    }
                    case 801: {
                        if (!g_state.slides.empty()) {
                            g_state.slides[g_state.currentSlide].bgColor = RGB(214,234,248);
                            InvalidateRect(g_hwnd, NULL, TRUE);
                        }
                        break;
                    }
                    case 802: {
                        if (!g_state.slides.empty()) {
                            g_state.slides[g_state.currentSlide].bgColor = RGB(44,62,80);
                            InvalidateRect(g_hwnd, NULL, TRUE);
                        }
                        break;
                    }
                }
                return;
            }
        }
    }
}

void HandleSlideClick(int mx, int my, int slideX, int slideY, int slideW, int slideH) {
    if (g_state.slides.empty()) return;
    auto& slide = g_state.slides[g_state.currentSlide];

    // Deselect all
    g_state.selectedElementId = -1;
    for (auto& e : slide.elements) e.isSelected = false;

    // Hit test elements in reverse (topmost first)
    for (int i = (int)slide.elements.size() - 1; i >= 0; i--) {
        auto& elem = slide.elements[i];
        int ex = slideX + (int)(elem.rect.left   * slideW / 100.0f);
        int ey = slideY + (int)(elem.rect.top    * slideH / 100.0f);
        int ew =          (int)((elem.rect.right  - elem.rect.left) * slideW / 100.0f);
        int eh =          (int)((elem.rect.bottom - elem.rect.top)  * slideH / 100.0f);

        if (mx >= ex && mx <= ex + ew && my >= ey && my <= ey + eh) {
            elem.isSelected = true;
            g_state.selectedElementId = elem.id;
            g_state.isDragging = true;
            g_state.dragStart = {mx, my};
            g_state.originalRect = elem.rect;
            break;
        }
    }
    InvalidateRect(g_hwnd, NULL, TRUE);
}

void HandleThumbnailClick(int mx, int my, int sidebarW, int /*winH*/) {
    int thumbW = sidebarW - 20;
    int thumbH = (int)(thumbW * 9.0f / 16.0f);
    int thumbX = 10;
    int thumbY = g_state.ribbonHeight + 28 - g_state.thumbnailScrollPos;

    for (int i = 0; i < (int)g_state.slides.size(); i++) {
        if (mx >= thumbX && mx <= thumbX + thumbW &&
            my >= thumbY && my <= thumbY + thumbH) {
            g_state.currentSlide = i;
            g_state.selectedElementId = -1;
            InvalidateRect(g_hwnd, NULL, TRUE);
            return;
        }
        thumbY += thumbH + 12;
    }
}

// ==================== TEXT EDITING ====================
void HandleKeyDown(WPARAM wParam) {
    if (g_state.selectedElementId < 0 || g_state.slides.empty()) return;

    for (auto& elem : g_state.slides[g_state.currentSlide].elements) {
        if (elem.id == g_state.selectedElementId) {
            if (wParam == VK_DELETE || wParam == VK_BACK) {
                if (!elem.text.empty()) {
                    SaveState();
                    elem.text.pop_back();
                    InvalidateRect(g_hwnd, NULL, TRUE);
                }
            }
            break;
        }
    }
}

void HandleChar(WPARAM wParam) {
    if (g_state.selectedElementId < 0 || g_state.slides.empty()) return;
    if (wParam < 32) return; // ignore control chars

    for (auto& elem : g_state.slides[g_state.currentSlide].elements) {
        if (elem.id == g_state.selectedElementId) {
            SaveState();
            elem.text += (wchar_t)wParam;
            InvalidateRect(g_hwnd, NULL, TRUE);
            break;
        }
    }
}

// ==================== DRAG ====================
void HandleMouseMove(int mx, int my, int winW, int winH) {
    bool redraw = false;

    // Ribbon hover
    if (g_state.activeTabIndex < (int)g_state.ribbonTabs.size()) {
        for (auto& group : g_state.ribbonTabs[g_state.activeTabIndex].groups) {
            for (auto& btn : group.buttons) {
                bool h = (mx >= btn.rect.left && mx < btn.rect.right &&
                          my >= btn.rect.top  && my < btn.rect.bottom);
                if (h != btn.isHovered) { btn.isHovered = h; redraw = true; }
            }
        }
    }

    // Drag selected element
    if (g_state.isDragging && g_state.selectedElementId >= 0 && !g_state.slides.empty()) {
        int slideX, slideY, slideW, slideH;
        GetSlideArea(winW, winH, slideX, slideY, slideW, slideH);

        int dx = (int)((mx - g_state.dragStart.x) * 100.0f / slideW);
        int dy = (int)((my - g_state.dragStart.y) * 100.0f / slideH);

        for (auto& elem : g_state.slides[g_state.currentSlide].elements) {
            if (elem.id == g_state.selectedElementId) {
                int ow = g_state.originalRect.right  - g_state.originalRect.left;
                int oh = g_state.originalRect.bottom - g_state.originalRect.top;
                int nl = max(0, min(95, g_state.originalRect.left + dx));
                int nt = max(0, min(95, g_state.originalRect.top  + dy));
                elem.rect = {nl, nt, nl + ow, nt + oh};
                redraw = true;
                break;
            }
        }
    }

    if (redraw) InvalidateRect(g_hwnd, NULL, TRUE);
}

// ==================== PAINT ====================
void OnPaint(HWND hwnd) {
    PAINTSTRUCT ps;
    HDC hdc = BeginPaint(hwnd, &ps);

    RECT clientRect;
    GetClientRect(hwnd, &clientRect);
    int W = clientRect.right, H = clientRect.bottom;

    // Double buffer
    HDC memDC = CreateCompatibleDC(hdc);
    HBITMAP memBmp = CreateCompatibleBitmap(hdc, W, H);
    HBITMAP oldBmp = (HBITMAP)SelectObject(memDC, memBmp);

    Graphics g(memDC);
    g.SetSmoothingMode(SmoothingModeAntiAlias);
    g.SetTextRenderingHint(TextRenderingHintClearTypeGridFit);

    // Main background
    SolidBrush mainBg(Color(255, 220, 224, 228));
    g.FillRectangle(&mainBg, 0, 0, W, H);

    // Draw panels
    DrawRibbon(g, W);
    DrawSidebar(g, W, H);
    DrawStatusBar(g, W, H);

    // Draw main slide
    if (!g_state.slides.empty()) {
        int slideX, slideY, slideW, slideH;
        GetSlideArea(W, H, slideX, slideY, slideW, slideH);

        // Drop shadow
        SolidBrush shadowBrush(Color(60, 0, 0, 0));
        g.FillRectangle(&shadowBrush, slideX + 6, slideY + 6, slideW, slideH);

        DrawSlide(g, slideX, slideY, slideW, slideH, g_state.slides[g_state.currentSlide], false);
    }

    // Color picker overlay
    if (g_state.colorPickerOpen)
        DrawColorPicker(g, W / 2 - 110, H / 2 - 80);

    BitBlt(hdc, 0, 0, W, H, memDC, 0, 0, SRCCOPY);
    SelectObject(memDC, oldBmp);
    DeleteObject(memBmp);
    DeleteDC(memDC);
    EndPaint(hwnd, &ps);
}

// ==================== WINDOW PROCEDURE ====================
LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    int mx = GET_X_LPARAM(lParam), my = GET_Y_LPARAM(lParam);

    switch (msg) {
        case WM_PAINT:
            OnPaint(hwnd);
            return 0;

        case WM_SIZE: {
            g_windowWidth  = LOWORD(lParam);
            g_windowHeight = HIWORD(lParam);
            g_scaleFactor = g_windowWidth / 1600.0f;
            if (g_scaleFactor < 0.5f) g_scaleFactor = 0.5f;
            InvalidateRect(hwnd, NULL, TRUE);
            return 0;
        }

        case WM_LBUTTONDOWN: {
            SetFocus(hwnd);
            SetCapture(hwnd);
            // Ribbon?
            if (my < g_state.ribbonHeight) {
                HandleRibbonClick(mx, my);
                return 0;
            }
            // Sidebar?
            if (mx < g_state.sidebarWidth) {
                HandleThumbnailClick(mx, my, g_state.sidebarWidth, g_windowHeight);
                return 0;
            }
            // Slide area
            {
                int slideX, slideY, slideW, slideH;
                GetSlideArea(g_windowWidth, g_windowHeight, slideX, slideY, slideW, slideH);
                if (mx >= slideX && mx <= slideX + slideW && my >= slideY && my <= slideY + slideH)
                    HandleSlideClick(mx, my, slideX, slideY, slideW, slideH);
                else {
                    // Click outside slide → deselect
                    if (!g_state.slides.empty())
                        for (auto& e : g_state.slides[g_state.currentSlide].elements)
                            e.isSelected = false;
                    g_state.selectedElementId = -1;
                    InvalidateRect(hwnd, NULL, TRUE);
                }
            }
            return 0;
        }

        case WM_LBUTTONUP:
            g_state.isDragging = false;
            ReleaseCapture();
            return 0;

        case WM_MOUSEMOVE:
            HandleMouseMove(mx, my, g_windowWidth, g_windowHeight);
            return 0;

        case WM_MOUSEWHEEL: {
            int delta = GET_WHEEL_DELTA_WPARAM(wParam);
            if (GetAsyncKeyState(VK_CONTROL) & 0x8000) {
                g_scaleFactor += delta > 0 ? 0.1f : -0.1f;
                g_scaleFactor = max(0.4f, min(3.0f, g_scaleFactor));
            } else {
                g_state.thumbnailScrollPos -= delta / 3;
                if (g_state.thumbnailScrollPos < 0) g_state.thumbnailScrollPos = 0;
            }
            InvalidateRect(hwnd, NULL, TRUE);
            return 0;
        }

        case WM_KEYDOWN:
            if (wParam == 'Z' && (GetAsyncKeyState(VK_CONTROL) & 0x8000)) { Undo(); return 0; }
            if (wParam == 'Y' && (GetAsyncKeyState(VK_CONTROL) & 0x8000)) { Redo(); return 0; }
            if (wParam == VK_DELETE && g_state.selectedElementId < 0) {
                DeleteCurrentSlide(); return 0;
            }
            HandleKeyDown(wParam);
            return 0;

        case WM_CHAR:
            HandleChar(wParam);
            return 0;

        case WM_DESTROY:
            PostQuitMessage(0);
            return 0;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

// ==================== ENTRY POINT ====================
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE, LPSTR, int nCmdShow) {
    // GDI+
    GdiplusStartupInput gdiplusInput;
    ULONG_PTR gdiplusToken;
    GdiplusStartup(&gdiplusToken, &gdiplusInput, NULL);

    CoInitialize(NULL);

    // Register window class
    WNDCLASSEX wc = {};
    wc.cbSize        = sizeof(WNDCLASSEX);
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc   = WndProc;
    wc.hInstance     = hInstance;
    wc.hCursor       = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.lpszClassName = L"PPTXEditorClass";
    wc.hIcon         = LoadIcon(NULL, IDI_APPLICATION);
    RegisterClassEx(&wc);

    // Create window
    HWND hwnd = CreateWindowEx(
        WS_EX_APPWINDOW,
        L"PPTXEditorClass",
        L"Presentation Editor — Professional",
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT,
        1280, 720,
        NULL, NULL, hInstance, NULL
    );
    g_hwnd = hwnd;

    // Init
    InitializeRibbon();
    InitializeDefaultSlides();
    SaveState(); // initial undo snapshot

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    // Message loop
    MSG msg = {};
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    GdiplusShutdown(gdiplusToken);
    CoUninitialize();
    return (int)msg.wParam;
}
