/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef, useMemo } from "react";
import { motion, AnimatePresence } from "motion/react";
import * as Icons from "lucide-react";
import pptxgen from "pptxgenjs";
import { initialSlides, Slide } from "./data/slides";

// Helper functions to handle flexible Slide formats
const getSlideMessages = (slide: Slide): string[] => {
  if (slide.keyMessages && slide.keyMessages.length > 0) {
    return slide.keyMessages;
  }
  if (typeof slide.content === "string") {
    return slide.content
      .split("\n")
      .map(item => item.replace(/^[•✦\-]\s*/, "").trim())
      .filter(item => item !== "");
  }
  if (Array.isArray(slide.content)) {
    return slide.content.map(item => item.replace(/^[•✦\-]\s*/, "").trim());
  }
  return [];
};

const getSlideLecture = (slide: Slide): string => {
  if (slide.lecture) return slide.lecture;
  if (slide.id <= 5) return "مقدمة عامة";
  if (slide.id <= 50) return "المحاضرة الأولى";
  return "المحاضرة الثانية";
};

const getSlideSection = (slide: Slide): string => {
  if (slide.section) return slide.section;
  if (slide.id <= 5) return "توجيهات المحاضر";
  if (slide.id <= 20) return "1. الافتتاح وكسر الجمود";
  if (slide.id <= 40) return "2. الذكاء الاصطناعي والعولمة";
  if (slide.id <= 50) return "3. حماية البيانات والإدمان";
  if (slide.id <= 70) return "1. إدارة الوقت والأولويات";
  if (slide.id <= 90) return "2. إدارة العادات والعلاقات";
  return "3. الانضباط واتخاذ القرارات";
};

const getSlideType = (slide: Slide): string => {
  if (slide.type) return slide.type;
  if (slide.isActivity) return "activity";
  if (slide.isCaseStudy) return "case_study";
  return "content";
};

const getSlideIcon = (slide: Slide): string => {
  if (slide.icon) return slide.icon;
  const type = getSlideType(slide);
  if (type === "activity") return "Smile";
  if (type === "case_study") return "FileText";
  if (type === "bible") return "Bookmark";
  if (type === "interactive") return "Cpu";
  return "Presentation";
};

export default function App() {
  // Application state
  const [slides, setSlides] = useState<Slide[]>(initialSlides);
  const [currentSlideIndex, setCurrentSlideIndex] = useState<number>(0);
  const [searchQuery, setSearchQuery] = useState<string>("");
  const [selectedLecture, setSelectedLecture] = useState<string>("الكل");
  const [selectedType, setSelectedType] = useState<string>("الكل");
  const [isPresenterMode, setIsPresenterMode] = useState<boolean>(false);
  const [isMuted, setIsMuted] = useState<boolean>(false);
  
  // Custom theme settings
  const [themeStyle, setThemeStyle] = useState<"cream" | "pure" | "gold">("cream");
  const [logoPosition, setLogoPosition] = useState<"bottom-right" | "top-left" | "bottom-left">("bottom-right");
  const [baseFontSize, setBaselineFontSize] = useState<"normal" | "large" | "extra">("large");

  // Timer / Stopwatch state for presentation tracking (120 minutes)
  const [timerSeconds, setTimerSeconds] = useState<number>(0);
  const [isTimerRunning, setIsTimerSecondsRunning] = useState<boolean>(false);

  // Doodle Drawing Board state
  const [isDrawMode, setIsDrawMode] = useState<boolean>(false);
  const [drawColor, setDrawColor] = useState<string>("#f59e0b"); // gold by default
  const [drawWidth, setDrawWidth] = useState<number>(4);
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const [isDrawing, setIsDrawing] = useState<boolean>(false);

  // Speaker notes live customizer state
  const [isEditingNotes, setIsEditingNotes] = useState<boolean>(false);
  const [tempNotes, setTempNotes] = useState<string>("");
  const [tempKeyMessages, setTempKeyMessages] = useState<string>("");
  const [tempTitle, setTempTitle] = useState<string>("");

  // Quiz interactive state
  const [quizAnswers, setQuizAnswers] = useState<Record<string, string | boolean>>({});
  const [quizSubmitted, setQuizAnswersSubmitted] = useState<Record<string, boolean>>({});
  const [showPrintIframeModal, setShowPrintIframeModal] = useState<boolean>(false);
  const [isFullscreen, setIsFullscreen] = useState<boolean>(false);

  // Eisenhower Matrix interactive state
  const [matrixTasks, setMatrixTasks] = useState<any[]>([
    { id: "1", text: "الصلاة والقراءة الروحية اليومية", category: "" },
    { id: "2", text: "تسليم مشروع العمل العاجل", category: "" },
    { id: "3", text: "تصفح ريلز إنستغرام لثلاث ساعات", category: "" },
    { id: "4", text: "مكالمة هاتفية طويلة غير مجدية", category: "" }
  ]);

  // Screen Time Calculator state
  const [screenTimeHours, setScreenTimeHours] = useState<number>(5);
  const [calculatedSavings, setCalculatedSavings] = useState<{ yearlyHours: number; lifetimeYears: number; potentialBooks: number }>({
    yearlyHours: 0,
    lifetimeYears: 0,
    potentialBooks: 0
  });

  // Custom 90-Day Plan Checklist and custom entry state
  const [checkedPlanHabits, setCheckedPlanHabits] = useState<Record<string, boolean>>({});
  const [customPlanHabitText, setCustomPlanHabitText] = useState<string>("");
  const [customHabitsList, setCustomHabitsList] = useState<string[]>([
    "صلاة وتأمل هادئ لمدة 10 دقائق صباحاً",
    "قراءة 15 صفحة من كتاب هادف يومياً",
    "صيام تكنولوجي أسبوعي (يوم كامل بدون تواصل سلبي)",
    "زيارة ومكالمة تشجيعية لشخص محتاج"
  ]);

  // Filtered slides list
  const filteredSlides = useMemo(() => {
    return slides.filter((slide) => {
      const lecture = getSlideLecture(slide);
      const section = getSlideSection(slide);
      const type = getSlideType(slide);
      const keyMessages = getSlideMessages(slide);

      const matchesSearch =
        slide.title.toLowerCase().includes(searchQuery.toLowerCase()) ||
        section.toLowerCase().includes(searchQuery.toLowerCase()) ||
        (slide.subtitle && slide.subtitle.toLowerCase().includes(searchQuery.toLowerCase())) ||
        keyMessages.some((msg) => msg.toLowerCase().includes(searchQuery.toLowerCase()));

      const matchesLecture =
        selectedLecture === "الكل" ||
        (selectedLecture === "مقدمات عامة" && lecture === "مقدمة عامة") ||
        (selectedLecture === "المحاضرة الأولى" && lecture === "المحاضرة الأولى") ||
        (selectedLecture === "المحاضرة الثانية" && lecture === "المحاضرة الثانية");

      const matchesType =
        selectedType === "الكل" ||
        (selectedType === "نشاط" && type === "activity") ||
        (selectedType === "آية" && type === "bible") ||
        (selectedType === "دراسة حالة" && type === "case_study") ||
        (selectedType === "تفاعلي" && type === "interactive") ||
        (selectedType === "محتوى" && (type === "content" || type === "intro" || type === "summary"));

      return matchesSearch && matchesLecture && matchesType;
    });
  }, [slides, searchQuery, selectedLecture, selectedType]);

  // Ensure current slide index is valid if filtered list changes
  const currentSlide = useMemo(() => {
    if (filteredSlides.length === 0) return null;
    const slide = filteredSlides[currentSlideIndex];
    if (!slide) {
      setCurrentSlideIndex(0);
      return filteredSlides[0];
    }
    return slide;
  }, [filteredSlides, currentSlideIndex]);

  // Sync edits when slide changes
  useEffect(() => {
    if (currentSlide) {
      setTempNotes(currentSlide.speakerNotes || "");
      const msgs = getSlideMessages(currentSlide);
      setTempKeyMessages(msgs.join("\n"));
      setTempTitle(currentSlide.title);
    }
    // Clear drawing canvas when slide changes to prevent stale annotations
    clearCanvas();
  }, [currentSlide]);

  // Timer interval hook
  useEffect(() => {
    let interval: any = null;
    if (isTimerRunning) {
      interval = setInterval(() => {
        setTimerSeconds((prev) => prev + 1);
      }, 1000);
    } else {
      clearInterval(interval);
    }
    return () => clearInterval(interval);
  }, [isTimerRunning]);

  // Screen Time Calculation effect
  useEffect(() => {
    const yearly = screenTimeHours * 365;
    // Over 50 years, total hours spent, converted to years (24 hours/day)
    const lifetimeHrs = screenTimeHours * 365 * 50;
    const lifetimeYrs = Number((lifetimeHrs / (24 * 365)).toFixed(1));
    // Assuming an average book takes 4 hours to read
    const books = Math.floor(yearly / 4);

    setCalculatedSavings({
      yearlyHours: yearly,
      lifetimeYears: lifetimeYrs,
      potentialBooks: books
    });
  }, [screenTimeHours]);

  // Format stopwatch helper
  const formatTime = (totalSeconds: number) => {
    const mins = Math.floor(totalSeconds / 60);
    const secs = totalSeconds % 60;
    return `${mins.toString().padStart(2, "0")}:${secs.toString().padStart(2, "0")}`;
  };

  // Keyboard navigation for slides
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.target instanceof HTMLInputElement || e.target instanceof HTMLTextAreaElement) {
        return; // Ignore key bindings inside inputs
      }
      if (e.key === "ArrowRight") {
        goToNextSlide();
      } else if (e.key === "ArrowLeft") {
        goToPrevSlide();
      } else if (e.key === "Space") {
        e.preventDefault();
        goToNextSlide();
      } else if (e.key === "Escape") {
        exitFullscreen();
      }
    };
    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [currentSlideIndex, filteredSlides, isFullscreen]);

  // Sync with native fullscreen exit (e.g. user presses Esc key on keyboard)
  useEffect(() => {
    const handleFullscreenChange = () => {
      const isCurrentlyNativeFullscreen = !!document.fullscreenElement;
      if (!isCurrentlyNativeFullscreen && isFullscreen) {
        setIsFullscreen(false);
      }
    };
    document.addEventListener("fullscreenchange", handleFullscreenChange);
    return () => document.removeEventListener("fullscreenchange", handleFullscreenChange);
  }, [isFullscreen]);

  const enterFullscreen = async () => {
    setIsFullscreen(true);
    try {
      const elem = document.documentElement;
      if (elem.requestFullscreen) {
        await elem.requestFullscreen();
      }
    } catch (err) {
      console.warn("Fullscreen request denied or not supported in iframe context:", err);
    }
  };

  const exitFullscreen = async () => {
    setIsFullscreen(false);
    try {
      if (document.fullscreenElement && document.exitFullscreen) {
        await document.exitFullscreen();
      }
    } catch (err) {
      console.warn("Exit fullscreen failed:", err);
    }
  };

  const goToNextSlide = () => {
    if (currentSlideIndex < filteredSlides.length - 1) {
      setCurrentSlideIndex((prev) => prev + 1);
    }
  };

  const goToPrevSlide = () => {
    if (currentSlideIndex > 0) {
      setCurrentSlideIndex((prev) => prev - 1);
    }
  };

  // Doodle drawing logic
  useEffect(() => {
    if (canvasRef.current) {
      const canvas = canvasRef.current;
      const ctx = canvas.getContext("2d");
      if (ctx) {
        ctx.lineCap = "round";
        ctx.lineJoin = "round";
      }
    }
  }, [isDrawMode]);

  const startDrawing = (e: React.MouseEvent<HTMLCanvasElement>) => {
    if (!isDrawMode || !canvasRef.current) return;
    const canvas = canvasRef.current;
    const ctx = canvas.getContext("2d");
    if (!ctx) return;

    // Get position relative to canvas bounds
    const rect = canvas.getBoundingClientRect();
    const x = ((e.clientX - rect.left) / rect.width) * canvas.width;
    const y = ((e.clientY - rect.top) / rect.height) * canvas.height;

    ctx.beginPath();
    ctx.moveTo(x, y);
    ctx.strokeStyle = drawColor;
    ctx.lineWidth = drawWidth;
    setIsDrawing(true);
  };

  const draw = (e: React.MouseEvent<HTMLCanvasElement>) => {
    if (!isDrawing || !isDrawMode || !canvasRef.current) return;
    const canvas = canvasRef.current;
    const ctx = canvas.getContext("2d");
    if (!ctx) return;

    const rect = canvas.getBoundingClientRect();
    const x = ((e.clientX - rect.left) / rect.width) * canvas.width;
    const y = ((e.clientY - rect.top) / rect.height) * canvas.height;

    ctx.lineTo(x, y);
    ctx.stroke();
  };

  const stopDrawing = () => {
    setIsDrawing(false);
  };

  const clearCanvas = () => {
    if (canvasRef.current) {
      const canvas = canvasRef.current;
      const ctx = canvas.getContext("2d");
      if (ctx) {
        ctx.clearRect(0, 0, canvas.width, canvas.height);
      }
    }
  };

  // Custom live slide edits saver
  const saveSlideEdits = () => {
    if (!currentSlide) return;
    const updated = slides.map((slide) => {
      if (slide.id === currentSlide.id) {
        return {
          ...slide,
          title: tempTitle,
          keyMessages: tempKeyMessages.split("\n").filter((line) => line.trim() !== ""),
          speakerNotes: tempNotes
        };
      }
      return slide;
    });
    setSlides(updated);
    setIsEditingNotes(false);
  };

  const resetAllSlides = () => {
    if (window.confirm("هل أنت متأكد من إعادة ضبط جميع الشرائح وتغييراتك إلى الوضع الأصلي الافتراضي للحقيبة؟")) {
      setSlides(initialSlides);
      setCurrentSlideIndex(0);
    }
  };

  // Handler for custom prints
  const handlePrint = () => {
    // Detect if inside an iframe sandbox
    const isInIframe = window.self !== window.top;
    if (isInIframe) {
      setShowPrintIframeModal(true);
    } else {
      window.print();
    }
  };

  // Handler for PowerPoint PPTX Exports
  const handleExportPPTX = () => {
    try {
      const pptx = new pptxgen();
      pptx.layout = "LAYOUT_16x9";

      // Set default text direction
      (pptx as any).rtl = true;

      const bgColors = {
        cream: "FBFBF9",
        pure: "FFFFFF",
        gold: "FFFDF5",
      };
      const currentBg = bgColors[themeStyle] || "FFFFFF";

      filteredSlides.forEach((slideItem, index) => {
        const pptxSlide = pptx.addSlide();
        pptxSlide.background = { fill: currentBg };

        // 1. Decorative custom color band on the right edge (as PPTX is RTL)
        // In PPTX, x=0 represents left edge, but since we are laying out RTL:
        // Let's add a nice decorative bar/border
        pptxSlide.addShape(pptx.ShapeType.rect, {
          x: 0,
          y: 0,
          w: 0.15,
          h: 5.625, // 16:9 standard height
          fill: { color: themeStyle === "pure" ? "2563EB" : "D97706" } // Blue vs Amber
        });

        // 2. Add Header (Lecture and Section Info)
        const lecture = getSlideLecture(slideItem);
        const section = getSlideSection(slideItem);
        const categoryText = `${lecture}   |   ${section}`;
        
        pptxSlide.addText(categoryText, {
          x: 0.5,
          y: 0.4,
          w: 9.0,
          h: 0.4,
          fontSize: 11,
          bold: true,
          color: themeStyle === "pure" ? "2563EB" : "D97706",
          align: "right",
          rtl: true,
          fontFace: "Arial"
        } as any);

        // 3. Add Slide Title
        pptxSlide.addText(slideItem.title, {
          x: 0.5,
          y: 0.8,
          w: 9.0,
          h: 0.8,
          fontSize: 24,
          bold: true,
          color: "1E293B", // slate-800
          align: "right",
          rtl: true,
          fontFace: "Arial"
        } as any);

        // 4. Add Subtitle if exists
        if (slideItem.subtitle) {
          pptxSlide.addText(slideItem.subtitle, {
            x: 0.5,
            y: 1.6,
            w: 9.0,
            h: 0.4,
            fontSize: 14,
            italic: true,
            color: "64748B",
            align: "right",
            rtl: true,
            fontFace: "Arial"
          } as any);
        }

        // 5. Add Main Messages list
        const messages = getSlideMessages(slideItem);
        let startY = slideItem.subtitle ? 2.1 : 1.7;
        
        if (messages && messages.length > 0) {
          messages.forEach((msg, msgIdx) => {
            pptxSlide.addText(`✦   ${msg}`, {
              x: 0.6,
              y: startY + (msgIdx * 0.7),
              w: 8.8,
              h: 0.6,
              fontSize: 14,
              color: "334155",
              align: "right",
              rtl: true,
              fontFace: "Arial"
            } as any);
          });
        }

        // 6. Brand Footer & Slide Index
        pptxSlide.addText(`GLOBAL FOR TRAINING   •   شريحة ${index + 1} من أصل ${filteredSlides.length}`, {
          x: 0.5,
          y: 5.1,
          w: 9.0,
          h: 0.3,
          fontSize: 10,
          color: "94A3B8",
          align: "center",
          rtl: true,
          fontFace: "Arial"
        } as any);

        // 7. Add actual Speaker Notes (accessible in PowerPoint Presenter View!)
        let notes = slideItem.speakerNotes || "";
        if (getSlideType(slideItem) === "activity") {
          notes = `[نشاط تفاعلي مباشر]\nالمطلوب: ${messages.join(" • ")}\n\n${notes}`;
        }
        if (notes) {
          pptxSlide.addNotes(notes);
        }
      });

      const fileName = `GLOBAL_Presentation_${selectedLecture === "الكل" ? "Full" : selectedLecture.replace(/\s+/g, "_")}.pptx`;
      pptx.writeFile({ fileName });
    } catch (err) {
      console.error("PPTX Generation Error: ", err);
      alert("حدث خطأ أثناء تصدير ملف PowerPoint. يرجى إعادة المحاولة.");
    }
  };

  // Helper dynamic icons renderer
  const renderDynamicIcon = (iconName: string, className: string = "w-6 h-6") => {
    const IconComponent = (Icons as any)[iconName];
    if (IconComponent) {
      return <IconComponent className={className} />;
    }
    return <Icons.Presentation className={className} />;
  };

  // Total percentages computation
  const getPlanProgress = () => {
    const checkedCount = Object.values(checkedPlanHabits).filter(Boolean).length;
    const totalCount = customHabitsList.length + 12; // default + extra
    return Math.round((checkedCount / totalCount) * 100) || 0;
  };

  return (
    <div className="min-h-screen bg-slate-100 text-slate-900 flex flex-col font-sans select-none overflow-x-hidden">
      
      {/* 1. Header Navigation Bar (no-print) */}
      <header className="no-print bg-white border-b border-slate-200 px-8 py-4 flex flex-col md:flex-row gap-4 items-center justify-between sticky top-0 z-50 shadow-sm">
        <div className="flex items-center gap-4">
          <div className="w-10 h-10 bg-blue-600 rounded-lg flex items-center justify-center text-white font-bold text-xl shadow-sm hover:scale-105 transition-transform">
            G
          </div>
          <div>
            <h1 className="text-xl font-bold font-display tracking-tight text-slate-800 flex items-center gap-2">
              GLOBAL FOR TRAINING <span className="text-blue-600 text-xs bg-blue-50 px-2.5 py-1 rounded-full border border-blue-200">ورشة شبابية</span>
            </h1>
            <p className="text-xs text-slate-400">حقيبة تدريبية تفاعلية حديثة ومتكاملة • أكثر من 100 شريحة احترافية قابلة للتعديل والطباعة</p>
          </div>
        </div>

        {/* Global Settings & Quick Utilities */}
        <div className="flex flex-wrap items-center gap-3">
          
          {/* Stopwatch / Lecture Timer */}
          <div className="flex items-center gap-2 bg-slate-50 px-3 py-1.5 rounded-xl border border-slate-200">
            <span className="text-xs text-slate-500 font-mono">عداد الورشة:</span>
            <span className="text-sm font-bold font-mono text-blue-600">{formatTime(timerSeconds)}</span>
            <button
              onClick={() => setIsTimerSecondsRunning(!isTimerRunning)}
              className="p-1 hover:text-blue-600 text-slate-400 transition-colors"
              title={isTimerRunning ? "إيقاف مؤقت" : "بدء العداد"}
            >
              {isTimerRunning ? <Icons.Pause className="w-3.5 h-3.5" /> : <Icons.Play className="w-3.5 h-3.5" />}
            </button>
            <button
              onClick={() => { setTimerSeconds(0); setIsTimerSecondsRunning(false); }}
              className="p-1 hover:text-red-500 text-slate-400 transition-colors"
              title="إعادة ضبط العداد"
            >
              <Icons.RotateCcw className="w-3.5 h-3.5" />
            </button>
          </div>

          {/* FontSize controls */}
          <div className="flex items-center gap-1 bg-slate-50 p-1 rounded-xl border border-slate-200">
            <span className="text-xs text-slate-500 px-2">الخط:</span>
            <button
              onClick={() => setBaselineFontSize("normal")}
              className={`text-xs px-2.5 py-1 rounded-lg transition-all ${baseFontSize === "normal" ? "bg-blue-600 text-white font-bold shadow-sm" : "text-slate-500 hover:text-slate-800"}`}
            >
              عادي
            </button>
            <button
              onClick={() => setBaselineFontSize("large")}
              className={`text-xs px-2.5 py-1 rounded-lg transition-all ${baseFontSize === "large" ? "bg-blue-600 text-white font-bold shadow-sm" : "text-slate-500 hover:text-slate-800"}`}
            >
              كبير
            </button>
            <button
              onClick={() => setBaselineFontSize("extra")}
              className={`text-xs px-2.5 py-1 rounded-lg transition-all ${baseFontSize === "extra" ? "bg-blue-600 text-white font-bold shadow-sm" : "text-slate-500 hover:text-slate-800"}`}
            >
              ضخم
            </button>
          </div>

          {/* Open in Full Page Button */}
          <button
            onClick={() => window.open(window.location.href, '_blank')}
            className="bg-slate-800 hover:bg-slate-900 text-white px-5 py-2 rounded-full text-sm font-bold font-display flex items-center gap-2 shadow-md shadow-slate-800/10 transition-all hover:scale-105"
            title="فتح الحقيبة في صفحة كاملة ومستقلة"
          >
            <Icons.ExternalLink className="w-4 h-4" />
            فتح في صفحة كاملة
          </button>

          {/* PDF Export Button */}
          <button
            onClick={handlePrint}
            className="bg-blue-600 hover:bg-blue-700 text-white px-5 py-2 rounded-full text-sm font-bold font-display flex items-center gap-2 shadow-md shadow-blue-600/10 transition-all hover:scale-105"
          >
            <Icons.Printer className="w-4 h-4" />
            تصدير PDF عالي الجودة
          </button>

          {/* PPTX Export Button */}
          <button
            onClick={handleExportPPTX}
            className="bg-amber-500 hover:bg-amber-600 text-slate-950 px-5 py-2 rounded-full text-sm font-bold font-display flex items-center gap-2 shadow-md shadow-amber-500/10 transition-all hover:scale-105"
            title="تصدير الشرائح كملف عرض تقديم ميكروسوفت باوربوينت تفاعلي"
          >
            <Icons.Download className="w-4 h-4" />
            تصدير PowerPoint PPTX
          </button>

          {/* Reset button */}
          <button
            onClick={resetAllSlides}
            className="border border-slate-200 hover:border-red-200 text-slate-400 hover:bg-red-50 hover:text-red-500 p-2.5 rounded-xl transition-all"
            title="إعادة ضبط الحقيبة للوضع الأصلي"
          >
            <Icons.Trash2 className="w-4 h-4" />
          </button>
        </div>
      </header>

      {/* 2. Top Navigation Filter Bar (no-print) */}
      <section className="no-print bg-slate-50 border-b border-slate-200 px-6 py-3 flex flex-wrap gap-4 items-center justify-between">
        {/* Lecture tab filters */}
        <div className="flex flex-wrap gap-1.5">
          {[
            { id: "الكل", label: "جميع الشرائح" },
            { id: "مقدمات عامة", label: "مقدمات منهجية" },
            { id: "المحاضرة الأولى", label: "المحاضرة 1: الذكاء الاصطناعي والعولمة" },
            { id: "المحاضرة الثانية", label: "المحاضرة 2: تطوير الذات وتجديد الذهن" }
          ].map((tab) => (
            <button
              key={tab.id}
              onClick={() => { setSelectedLecture(tab.id); setCurrentSlideIndex(0); }}
              className={`px-4 py-1.5 rounded-full text-xs font-bold font-display transition-all ${selectedLecture === tab.id ? "bg-blue-600 text-white shadow-sm shadow-blue-600/10" : "bg-white text-slate-600 hover:text-slate-900 border border-slate-200"}`}
            >
              {tab.label}
            </button>
          ))}
        </div>

        {/* Slide category tag filters */}
        <div className="flex items-center gap-1.5 flex-wrap">
          <span className="text-xs text-slate-500 ml-1.5">تصنيف المحتوى:</span>
          {[
            { id: "الكل", label: "الكل" },
            { id: "محتوى", label: "مادة نظرية" },
            { id: "نشاط", label: "أنشطة تفاعلية" },
            { id: "آية", label: "آيات وحفظ" },
            { id: "تفاعلي", label: "تطبيقات حاسبة" },
            { id: "دراسة حالة", label: "دراسات حالة" }
          ].map((typeTab) => (
            <button
              key={typeTab.id}
              onClick={() => { setSelectedType(typeTab.id); setCurrentSlideIndex(0); }}
              className={`px-3 py-1 rounded-full text-xs transition-all ${selectedType === typeTab.id ? "bg-slate-200 text-slate-800 font-bold" : "text-slate-500 hover:text-slate-800 hover:bg-slate-100"}`}
            >
              {typeTab.label}
            </button>
          ))}
        </div>

        {/* Interactive Search query */}
        <div className="relative w-full md:w-64">
          <span className="absolute inset-y-0 right-0 pr-3 flex items-center text-slate-400">
            <Icons.Search className="w-4 h-4" />
          </span>
          <input
            type="text"
            placeholder="ابحث بالشرائح والأهداف..."
            value={searchQuery}
            onChange={(e) => { setSearchQuery(e.target.value); setCurrentSlideIndex(0); }}
            className="w-full bg-white text-xs text-slate-800 pr-9 pl-4 py-2 rounded-full border border-slate-200 focus:outline-none focus:border-blue-500/50 transition-all placeholder:text-slate-400"
          />
          {searchQuery && (
            <button onClick={() => setSearchQuery("")} className="absolute inset-y-0 left-0 pl-3 flex items-center text-slate-400 hover:text-slate-600">
              <Icons.X className="w-3.5 h-3.5" />
            </button>
          )}
        </div>
      </section>

      {/* 3. Main Split View Workspace (no-print) */}
      <main className="no-print flex-1 flex flex-col lg:flex-row overflow-hidden">
        
        {/* Left-Sidebar: Slides Outline & Index (RTL scrollable) */}
        <aside className="w-full lg:w-80 bg-white border-l border-slate-200 flex flex-col justify-between overflow-y-auto max-h-[400px] lg:max-h-none order-2 lg:order-1">
          <div className="p-4 border-b border-slate-100 flex items-center justify-between bg-slate-50">
            <h3 className="text-xs font-bold uppercase tracking-wider text-slate-500 flex items-center gap-1.5">
              <Icons.List className="w-4 h-4 text-blue-600" />
              مخطط الحقيبة ({filteredSlides.length} شريحة)
            </h3>
            <span className="text-[10px] bg-white px-2.5 py-0.5 rounded-full border border-slate-200 text-slate-500 font-bold">
              {selectedLecture}
            </span>
          </div>

          {/* Slides List Grid */}
          <div className="flex-1 divide-y divide-slate-100 overflow-y-auto">
            {filteredSlides.length === 0 ? (
              <div className="p-8 text-center text-slate-400 text-xs">
                <Icons.FileQuestion className="w-10 h-10 mx-auto mb-2 text-slate-300" />
                لا توجد شرائح تطابق الفلاتر الحالية
              </div>
            ) : (
              filteredSlides.map((slide, idx) => {
                const isSelected = currentSlideIndex === idx;
                return (
                  <button
                    key={slide.id}
                    onClick={() => setCurrentSlideIndex(idx)}
                    className={`w-full text-right p-3.5 flex gap-3 transition-all ${isSelected ? "bg-blue-50/70 border-r-4 border-blue-600 shadow-sm" : "hover:bg-slate-50/60"}`}
                  >
                    <div className={`mt-0.5 w-6 h-6 rounded-lg flex items-center justify-center text-[10px] font-bold transition-all ${isSelected ? "bg-blue-600 text-white shadow-sm" : "bg-slate-100 text-slate-500 border border-slate-200"}`}>
                      {slide.id}
                    </div>
                    <div className="flex-1 min-w-0">
                      <div className="flex items-center justify-between gap-1 mb-0.5">
                        <span className={`text-[10px] font-mono truncate ${isSelected ? "text-blue-600 font-bold" : "text-slate-400"}`}>
                          {getSlideSection(slide)}
                        </span>
                        <span className="text-[10px] text-slate-400 shrink-0">
                          {getSlideType(slide) === "activity" && "🎯 نشاط"}
                          {getSlideType(slide) === "bible" && "📖 آية"}
                          {getSlideType(slide) === "case_study" && "📂 حالة"}
                          {getSlideType(slide) === "interactive" && "⚙️ تفاعلي"}
                        </span>
                      </div>
                      <h4 className={`text-xs font-bold truncate ${isSelected ? "text-slate-900" : "text-slate-600 hover:text-slate-900"}`}>
                        {slide.title}
                      </h4>
                    </div>
                  </button>
                );
              })
            )}
          </div>

          <div className="p-4 bg-slate-50 border-t border-slate-200 text-[10px] text-slate-400 flex justify-between items-center">
            <span>حقوق الطبع محفوظة 2026 ©</span>
            <span className="font-bold text-blue-600">GLOBAL FOR TRAINING</span>
          </div>
        </aside>

        {/* Central Presentation Canvas Frame */}
        <section className="flex-1 bg-slate-100 p-6 flex flex-col gap-6 items-center justify-start overflow-y-auto order-1 lg:order-2">
          
          {/* Controls above stage */}
          <div className="w-full max-w-4xl flex items-center justify-between gap-4">
            <div className="flex items-center gap-2">
              <button
                onClick={goToPrevSlide}
                disabled={currentSlideIndex === 0}
                className="bg-white hover:bg-slate-50 disabled:opacity-30 disabled:hover:bg-white border border-slate-200 px-3.5 py-1.5 rounded-full text-xs font-bold flex items-center gap-1.5 transition-all text-slate-700 shadow-sm"
              >
                <Icons.ArrowRight className="w-4 h-4" />
                السابقة
              </button>
              
              <span className="text-xs text-slate-500 font-mono">
                الشريحة <strong className="text-slate-900">{currentSlideIndex + 1}</strong> من أصل <strong className="text-slate-500">{filteredSlides.length}</strong>
              </span>

              <button
                onClick={goToNextSlide}
                disabled={currentSlideIndex === filteredSlides.length - 1}
                className="bg-white hover:bg-slate-50 disabled:opacity-30 disabled:hover:bg-white border border-slate-200 px-3.5 py-1.5 rounded-full text-xs font-bold flex items-center gap-1.5 transition-all text-slate-700 shadow-sm"
              >
                التالية
                <Icons.ArrowLeft className="w-4 h-4" />
              </button>
            </div>

            {/* Custom Theme selection tabs */}
            <div className="flex items-center gap-2">
              <span className="text-[10px] text-slate-500">النمط البصري:</span>
              <div className="flex bg-white p-1 rounded-full border border-slate-200 shadow-sm">
                <button
                  onClick={() => setThemeStyle("cream")}
                  className={`text-[10px] px-3 py-1 rounded-full transition-all ${themeStyle === "cream" ? "bg-amber-100 text-amber-900 font-bold" : "text-slate-500 hover:text-slate-800"}`}
                >
                  عاجي كلاسيك
                </button>
                <button
                  onClick={() => setThemeStyle("pure")}
                  className={`text-[10px] px-3 py-1 rounded-full transition-all ${themeStyle === "pure" ? "bg-blue-50 text-blue-600 font-bold" : "text-slate-500 hover:text-slate-800"}`}
                >
                  أبيض ناصع
                </button>
                <button
                  onClick={() => setThemeStyle("gold")}
                  className={`text-[10px] px-3 py-1 rounded-full transition-all ${themeStyle === "gold" ? "bg-amber-500 text-slate-950 font-bold" : "text-slate-500 hover:text-slate-800"}`}
                >
                  ذهبي ملكي
                </button>
              </div>
            </div>

            <div className="flex items-center gap-2">
              {/* Virtual Laser / Marker whiteboard mode */}
              <button
                onClick={() => { setIsDrawMode(!isDrawMode); clearCanvas(); }}
                className={`px-3 py-1.5 rounded-full text-xs font-bold flex items-center gap-1.5 transition-all ${isDrawMode ? "bg-blue-600 text-white shadow-md shadow-blue-600/10" : "bg-white hover:bg-slate-50 text-slate-600 border border-slate-200 shadow-sm"}`}
                title="تفعيل سبورة الرسم والتظليل التفاعلي فوق الشريحة"
              >
                <Icons.Edit3 className="w-3.5 h-3.5" />
                {isDrawMode ? "تعطيل الرسم" : "قلم التظليل"}
              </button>
              
              {isDrawMode && (
                <button
                  onClick={clearCanvas}
                  className="bg-red-50 hover:bg-red-100 text-red-600 px-3 py-1.5 rounded-full text-xs border border-red-200 transition-all"
                >
                  مسح الرسومات
                </button>
              )}

              {/* Fullscreen presentation button */}
              <button
                onClick={enterFullscreen}
                className="bg-blue-600 hover:bg-blue-700 text-white px-3.5 py-1.5 rounded-full text-xs font-bold flex items-center gap-1.5 transition-all shadow-md shadow-blue-600/15 hover:scale-105"
                title="بدء عرض الشرائح بملء الشاشة بالكامل وبدون حواف"
              >
                <Icons.Maximize2 className="w-3.5 h-3.5" />
                عرض ملء الشاشة 📺
              </button>
            </div>
          </div>

          {/* Actual 16:9 Slide Widescreen Presentation Canvas */}
          <div className={isFullscreen 
            ? "fixed inset-0 w-screen h-screen z-[9999] bg-slate-950 flex flex-col items-center justify-center p-4 md:p-12 overflow-hidden" 
            : "contents"
          }>
            <div className={isFullscreen 
              ? "w-full max-w-6xl aspect-video relative rounded-2xl overflow-hidden shadow-2xl border border-slate-800 bg-white"
              : "w-full max-w-4xl aspect-video relative rounded-3xl overflow-hidden shadow-2xl border border-slate-200 bg-white"
            }>
              
              {/* Interactive Draw Layer */}
              {isDrawMode && (
                <canvas
                  ref={canvasRef}
                  width={1280}
                  height={720}
                  onMouseDown={startDrawing}
                  onMouseMove={draw}
                  onMouseUp={stopDrawing}
                  onMouseLeave={stopDrawing}
                  className="absolute inset-0 z-40 cursor-crosshair w-full h-full pointer-events-auto"
                />
              )}

            {/* Slide Content Layout with Animated Presence */}
            <AnimatePresence mode="wait">
              {currentSlide && (
                <motion.div
                  key={currentSlide.id}
                  initial={{ opacity: 0, x: 20 }}
                  animate={{ opacity: 1, x: 0 }}
                  exit={{ opacity: 0, x: -20 }}
                  transition={{ duration: 0.3 }}
                  dir="rtl"
                  className={`absolute inset-0 w-full h-full p-8 md:p-12 flex flex-col justify-between transition-all select-text select-all ${
                    themeStyle === "cream" ? "bg-[#fbfbf9]" : 
                    themeStyle === "pure" ? "bg-white" : 
                    "bg-[#FAF7EE] border-t-8 border-amber-500"
                  } text-slate-900`}
                >
                  
                  {/* Decorative luxury elements (thin border lines, geometric background lights) */}
                  <div className={`absolute inset-4 border rounded-2xl pointer-events-none ${
                    themeStyle === "pure" ? "border-blue-500/10" : "border-amber-500/10"
                  }`} />
                  
                  {/* Slide Top bar */}
                  <div className="flex items-start justify-between z-10">
                    <div className="flex items-center gap-3">
                      <div className={`p-2.5 rounded-xl border ${
                        themeStyle === "pure" 
                          ? "bg-blue-50 text-blue-600 border-blue-200" 
                          : "bg-amber-500/10 text-amber-600 border-amber-500/20"
                      }`}>
                        {renderDynamicIcon(getSlideIcon(currentSlide), "w-6 h-6")}
                      </div>
                      <div>
                        <span className={`text-xs font-bold tracking-wider ${
                          themeStyle === "pure" ? "text-blue-600" : "text-amber-600"
                        }`}>
                          {getSlideLecture(currentSlide)} • {getSlideSection(currentSlide)}
                        </span>
                        <h2 className="text-base font-bold text-slate-600 mt-0.5">
                          {currentSlide.subtitle || "حقيبة متميزة لتأهيل الشباب المبتكر"}
                        </h2>
                      </div>
                    </div>

                    {/* Top corner logo option */}
                    {logoPosition === "top-left" && (
                      <div className="bg-white text-slate-800 border border-slate-200 px-3.5 py-1 rounded-xl text-center shadow-sm">
                        <div className="text-[10px] tracking-widest text-blue-600 font-black font-mono">GLOBAL</div>
                        <div className="text-[8px] text-slate-400 font-bold">FOR TRAINING</div>
                      </div>
                    )}
                  </div>

                  {/* Slide Core Content Body */}
                  <div className="my-auto z-10 flex flex-col gap-6 justify-center">
                    
                    {/* Slide Dynamic Title with FontSize */}
                    <h3 className={`font-bold font-display tracking-tight text-slate-950 ${
                      baseFontSize === "normal" ? "text-2xl md:text-3xl" :
                      baseFontSize === "large" ? "text-3xl md:text-4xl leading-tight" :
                      "text-4xl md:text-5xl leading-tight"
                    }`}>
                      {currentSlide.title}
                    </h3>

                    {/* Slide Bullet points or core textual details */}
                    <div className="flex flex-col gap-4 pl-4">
                      {getSlideMessages(currentSlide).map((msg, index) => (
                        <div key={index} className="flex items-start gap-3">
                          <span className={`mt-1.5 text-sm shrink-0 ${
                            themeStyle === "pure" ? "text-blue-500" : "text-amber-500"
                          }`}>✦</span>
                          <p className={`font-medium text-slate-800 ${
                            baseFontSize === "normal" ? "text-sm md:text-base" :
                            baseFontSize === "large" ? "text-base md:text-lg" :
                            "text-lg md:text-xl"
                          }`}>
                            {msg}
                          </p>
                        </div>
                      ))}
                    </div>

                    {/* DYNAMIC COMPONENT: Live Quiz & interactive elements */}
                    {currentSlide.interactiveConfig && (
                      <div className={`mt-4 p-4 rounded-2xl border transition-all ${
                        themeStyle === "pure" 
                          ? "bg-blue-50/40 border-blue-100" 
                          : "bg-amber-500/5 border-amber-500/10"
                      }`}>
                        
                        {/* Interactive Case 1: True / False interactive Battle */}
                        {currentSlide.interactiveConfig.quizType === "input_calculation" && (
                          <div className="flex flex-col gap-3">
                            <span className={`text-xs font-bold block ${
                              themeStyle === "pure" ? "text-blue-600" : "text-amber-600"
                            }`}>🧮 حاسبة وقت الشاشة المستهلك وتأثيره على المستقبل:</span>
                            <div className="flex items-center gap-4">
                              <input
                                type="range"
                                min="1"
                                max="15"
                                value={screenTimeHours}
                                onChange={(e) => setScreenTimeHours(Number(e.target.value))}
                                className={`w-1/2 cursor-pointer ${
                                  themeStyle === "pure" ? "accent-blue-600" : "accent-amber-500"
                                }`}
                              />
                              <span className={`text-sm font-bold bg-white px-3 py-1 rounded-lg border ${
                                themeStyle === "pure" ? "text-blue-600 border-blue-200" : "text-slate-900 border-slate-200"
                              }`}>
                                {screenTimeHours} ساعات يومياً
                              </span>
                            </div>
                            <div className="grid grid-cols-3 gap-3 text-center mt-2">
                              <div className="bg-white p-2 rounded-xl border border-slate-100">
                                <span className="text-[10px] text-slate-500 block">ساعاتك السنوية</span>
                                <span className={`text-sm font-black font-mono ${
                                  themeStyle === "pure" ? "text-blue-600" : "text-amber-600"
                                }`}>{calculatedSavings.yearlyHours} ساعة</span>
                              </div>
                              <div className="bg-white p-2 rounded-xl border border-slate-100">
                                <span className="text-[10px] text-slate-500 block">من عمرك (خلال 50 سنة)</span>
                                <span className="text-sm font-black text-red-600 font-mono">{calculatedSavings.lifetimeYears} سنة كاملة!</span>
                              </div>
                              <div className="bg-white p-2 rounded-xl border border-slate-100">
                                <span className="text-[10px] text-slate-500 block">مكافئ كتب تقرأها</span>
                                <span className="text-sm font-black text-emerald-600 font-mono">{calculatedSavings.potentialBooks} كتاب</span>
                              </div>
                            </div>
                          </div>
                        )}

                        {/* Interactive Case 2: Priority Grid categorizer */}
                        {currentSlide.interactiveConfig.quizType === "matrix" && (
                          <div className="flex flex-col gap-3">
                            <span className={`text-xs font-bold block ${
                              themeStyle === "pure" ? "text-blue-600" : "text-amber-600"
                            }`}>📂 صنف المهام التالية في مصفوفة الأولويات:</span>
                            <div className="grid grid-cols-2 gap-2">
                              {matrixTasks.map((task) => (
                                <div key={task.id} className="bg-white p-2.5 rounded-xl border border-slate-100 flex items-center justify-between text-xs">
                                  <span className="font-bold text-slate-800">{task.text}</span>
                                  <select
                                    value={task.category}
                                    onChange={(e) => {
                                      const updated = matrixTasks.map(t => t.id === task.id ? { ...t, category: e.target.value } : t);
                                      setMatrixTasks(updated);
                                    }}
                                    className={`rounded px-1.5 py-0.5 border focus:outline-none text-[11px] ${
                                      themeStyle === "pure" 
                                        ? "bg-blue-50 text-blue-800 border-blue-200" 
                                        : "bg-amber-500/10 text-amber-800 border-amber-500/20"
                                    }`}
                                  >
                                    <option value="">صنف المربع</option>
                                    <option value="urgent_important">عاجل ومهم</option>
                                    <option value="not_urgent_important">غير عاجل ومهم</option>
                                    <option value="urgent_not_important">عاجل وغير مهم</option>
                                    <option value="not_urgent_not_important">غير عاجل وغير مهم</option>
                                  </select>
                                </div>
                              ))}
                            </div>
                            <div className="grid grid-cols-2 gap-2 mt-2">
                              <div className="bg-red-50 p-2 rounded-xl border border-red-100 text-center">
                                <span className="text-[10px] text-red-700 font-bold block">1. عاجل ومهم (افعل الآن)</span>
                                <p className="text-[9px] text-slate-600 mt-1">
                                  {matrixTasks.filter(t => t.category === "urgent_important").map(t => t.text).join(" • ") || "لا توجد مهام"}
                                </p>
                              </div>
                              <div className="bg-emerald-50 p-2 rounded-xl border border-emerald-100 text-center">
                                <span className="text-[10px] text-emerald-700 font-bold block">2. غير عاجل ومهم (خطط وركز)</span>
                                <p className="text-[9px] text-slate-600 mt-1">
                                  {matrixTasks.filter(t => t.category === "not_urgent_important").map(t => t.text).join(" • ") || "لا توجد مهام"}
                                </p>
                              </div>
                            </div>
                          </div>
                        )}

                        {/* Interactive Case 3: 90-Day habit planner customizer */}
                        {currentSlide.interactiveConfig.quizType === "plan_grid" && (
                          <div className="flex flex-col gap-3">
                            <span className={`text-xs font-bold block ${
                              themeStyle === "pure" ? "text-blue-600" : "text-amber-600"
                            }`}>🗓️ صمم جدول عاداتك الـ 90 يوماً بنفسك واطبع خطتك:</span>
                            <div className="flex gap-2">
                              <input
                                type="text"
                                placeholder="اكتب عادة شخصية تود الالتزام بها..."
                                value={customPlanHabitText}
                                onChange={(e) => setCustomPlanHabitText(e.target.value)}
                                className={`flex-1 bg-white text-xs px-3 py-1.5 rounded-xl border border-slate-200 text-slate-900 focus:outline-none ${
                                  themeStyle === "pure" ? "focus:border-blue-500" : "focus:border-amber-500"
                                }`}
                              />
                              <button
                                onClick={() => {
                                  if (customPlanHabitText.trim()) {
                                    setCustomHabitsList([...customHabitsList, customPlanHabitText.trim()]);
                                    setCustomPlanHabitText("");
                                  }
                                }}
                                className={`font-bold text-xs px-4 py-1.5 rounded-xl transition-all ${
                                  themeStyle === "pure"
                                    ? "bg-blue-600 hover:bg-blue-700 text-white shadow-sm"
                                    : "bg-amber-500 hover:bg-amber-600 text-slate-950"
                                }`}
                              >
                                إضافة
                              </button>
                            </div>
                            
                            <div className="grid grid-cols-2 gap-1.5 max-h-[140px] overflow-y-auto mt-1 p-1 bg-white rounded-xl border border-slate-100">
                              {customHabitsList.map((habit, idx) => (
                                <label key={idx} className="flex items-center gap-2 text-[10px] text-slate-800 cursor-pointer hover:bg-slate-50 p-1 rounded transition-all">
                                  <input
                                    type="checkbox"
                                    checked={!!checkedPlanHabits[habit]}
                                    onChange={(e) => setCheckedPlanHabits({ ...checkedPlanHabits, [habit]: e.target.checked })}
                                    className={themeStyle === "pure" ? "accent-blue-600" : "accent-amber-500"}
                                  />
                                  <span className={checkedPlanHabits[habit] ? "line-through text-slate-400" : ""}>{habit}</span>
                                </label>
                              ))}
                            </div>

                            <div className="flex items-center justify-between text-xs mt-1">
                              <span className="text-slate-500">معدل التزامك المتوقع:</span>
                              <span className={`font-bold font-mono ${
                                themeStyle === "pure" ? "text-blue-600" : "text-amber-600"
                              }`}>{getPlanProgress()}% من الأهداف المحددة</span>
                            </div>
                          </div>
                        )}

                      </div>
                    )}
                  </div>

                  {/* Slide Bottom Bar with GLOBAL FOR TRAINING badge */}
                  <div className="flex items-center justify-between mt-auto pt-4 border-t border-slate-200/60 z-10">
                    <span className="text-xs font-bold font-mono text-slate-400">
                      دليل المحاضر التفاعلي لعام 2026 • ورشة عمل الشباب
                    </span>

                    {/* Integrated logo */}
                    {logoPosition === "bottom-right" && (
                      <div className="flex items-center gap-2 bg-slate-900 text-white border border-slate-800 px-4 py-1.5 rounded-full shadow-lg">
                        <div className={`w-2.5 h-2.5 rounded-full animate-pulse ${
                          themeStyle === "pure" ? "bg-blue-500" : "bg-amber-500"
                        }`} />
                        <span className={`text-[10px] tracking-widest font-black font-mono ${
                          themeStyle === "pure" ? "text-blue-400" : "text-amber-400"
                        }`}>GLOBAL FOR TRAINING</span>
                      </div>
                    )}
                  </div>

                </motion.div>
              )}
            </AnimatePresence>
          </div>

          {/* FULLSCREEN FLOAT CONTROLS HUD */}
          {isFullscreen && (
            <div dir="rtl" className="mt-6 flex flex-col md:flex-row items-center justify-between gap-6 w-full max-w-6xl bg-slate-900/95 border border-slate-800 px-6 py-4 rounded-3xl text-white shadow-2xl z-50 animate-fade-in backdrop-blur-md">
              <div className="flex items-center gap-3">
                <button
                  onClick={goToPrevSlide}
                  disabled={currentSlideIndex === 0}
                  className="bg-slate-800 hover:bg-slate-700 disabled:opacity-30 border border-slate-700 px-4 py-2 rounded-xl text-xs font-bold flex items-center gap-2 transition-all"
                >
                  <Icons.ArrowRight className="w-4 h-4 text-amber-500" />
                  السابقة
                </button>
                <span className="text-xs text-slate-300 font-mono">
                  الشريحة <strong className="text-amber-500 text-sm font-bold">{currentSlideIndex + 1}</strong> من <strong className="text-slate-400">{filteredSlides.length}</strong>
                </span>
                <button
                  onClick={goToNextSlide}
                  disabled={currentSlideIndex === filteredSlides.length - 1}
                  className="bg-slate-800 hover:bg-slate-700 disabled:opacity-30 border border-slate-700 px-4 py-2 rounded-xl text-xs font-bold flex items-center gap-2 transition-all"
                >
                  التالية
                  <Icons.ArrowLeft className="w-4 h-4 text-amber-500" />
                </button>
              </div>

              {/* Center controls: whiteboard drawing & visual theme switcher */}
              <div className="flex items-center gap-4 flex-wrap justify-center">
                {/* Visual theme switcher */}
                <div className="flex items-center gap-2">
                  <span className="text-[10px] text-slate-400">النمط:</span>
                  <div className="flex bg-slate-950 p-1 rounded-xl border border-slate-800">
                    <button
                      onClick={() => setThemeStyle("cream")}
                      className={`text-[10px] px-3 py-1.5 rounded-lg transition-all ${themeStyle === "cream" ? "bg-amber-500 text-slate-950 font-bold" : "text-slate-400 hover:text-white"}`}
                    >
                      عاجي
                    </button>
                    <button
                      onClick={() => setThemeStyle("pure")}
                      className={`text-[10px] px-3 py-1.5 rounded-lg transition-all ${themeStyle === "pure" ? "bg-blue-600 text-white font-bold" : "text-slate-400 hover:text-white"}`}
                    >
                      أبيض
                    </button>
                    <button
                      onClick={() => setThemeStyle("gold")}
                      className={`text-[10px] px-3 py-1.5 rounded-lg transition-all ${themeStyle === "gold" ? "bg-amber-500 text-slate-950 font-bold" : "text-slate-400 hover:text-white"}`}
                    >
                      ذهبي
                    </button>
                  </div>
                </div>

                {/* Drawing mode */}
                <button
                  onClick={() => { setIsDrawMode(!isDrawMode); clearCanvas(); }}
                  className={`px-3 py-1.5 rounded-xl text-xs font-bold flex items-center gap-1.5 transition-all ${isDrawMode ? "bg-blue-600 text-white" : "bg-slate-800 hover:bg-slate-700 text-slate-300"}`}
                >
                  <Icons.Edit3 className="w-3.5 h-3.5" />
                  {isDrawMode ? "تعطيل القلم" : "قلم التظليل"}
                </button>

                {isDrawMode && (
                  <button
                    onClick={clearCanvas}
                    className="bg-red-950 hover:bg-red-900 text-red-400 px-3 py-1.5 rounded-xl text-xs border border-red-900 transition-all"
                  >
                    مسح الرسم
                  </button>
                )}
              </div>

              {/* Right controls: exit */}
              <button
                onClick={exitFullscreen}
                className="bg-red-600 hover:bg-red-700 text-white px-5 py-2.5 rounded-xl text-xs font-bold flex items-center gap-1.5 transition-all hover:scale-105"
              >
                <Icons.Minimize2 className="w-4 h-4" />
                إنهاء العرض (Esc)
              </button>
            </div>
          )}
        </div>

          {/* 4. Bottom Guides & Lecturer Notes Board (no-print) */}
          <div className="w-full max-w-4xl bg-white rounded-3xl border border-slate-200 shadow-md p-6 flex flex-col md:flex-row gap-6 justify-between items-start">
            
            {/* Lecturer notes left block */}
            <div className="flex-1 flex flex-col gap-3 w-full">
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <span className="text-xs font-bold uppercase tracking-wider text-slate-500 flex items-center gap-1.5">
                  <Icons.BookOpen className="w-4 h-4 text-blue-600" />
                  ملاحظات وتوجيهات المحاضر (محتوى الحقيبة الأصلي)
                </span>
                
                <button
                  onClick={() => setIsEditingNotes(!isEditingNotes)}
                  className="text-xs font-bold text-blue-600 hover:text-blue-750 flex items-center gap-1"
                >
                  <Icons.Edit className="w-3.5 h-3.5" />
                  {isEditingNotes ? "عرض الملاحظات" : "تعديل الملاحظة"}
                </button>
              </div>

              {isEditingNotes ? (
                <div className="flex flex-col gap-3 w-full">
                  <div>
                    <label className="text-[10px] text-slate-500 mb-1 block">عنوان الشريحة:</label>
                    <input
                      type="text"
                      value={tempTitle}
                      onChange={(e) => setTempTitle(e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 focus:border-blue-500/50 focus:outline-none rounded-xl px-3 py-1.5 text-xs text-slate-800"
                    />
                  </div>
                  <div>
                    <label className="text-[10px] text-slate-500 mb-1 block">النقاط الرئيسية (شريحة لكل سطر):</label>
                    <textarea
                      value={tempKeyMessages}
                      onChange={(e) => setTempKeyMessages(e.target.value)}
                      rows={3}
                      className="w-full bg-slate-50 border border-slate-200 focus:border-blue-500/50 focus:outline-none rounded-xl px-3 py-1.5 text-xs text-slate-800"
                    />
                  </div>
                  <div>
                    <label className="text-[10px] text-slate-500 mb-1 block">كلام المحاضر وتوجيه الأنشطة:</label>
                    <textarea
                      value={tempNotes}
                      onChange={(e) => setTempNotes(e.target.value)}
                      rows={4}
                      className="w-full bg-slate-50 border border-slate-200 focus:border-blue-500/50 focus:outline-none rounded-xl px-3 py-1.5 text-xs text-slate-800"
                    />
                  </div>
                  <div className="flex justify-end gap-2 mt-1">
                    <button onClick={() => setIsEditingNotes(false)} className="text-xs text-slate-400 px-3 py-1.5 rounded-lg hover:bg-slate-100">
                      إلغاء
                    </button>
                    <button onClick={saveSlideEdits} className="bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold px-4 py-1.5 rounded-xl transition-all shadow-sm">
                      حفظ التعديلات الحالية
                    </button>
                  </div>
                </div>
              ) : (
                <div className="text-xs text-slate-600 leading-relaxed font-sans rtl">
                  <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100 mb-2">
                    <strong className="text-blue-600 block mb-1">📢 كيف تقدم هذه الشريحة؟</strong>
                    <p className="text-slate-600 font-sans">
                      {currentSlide?.speakerNotes || (currentSlide && getSlideType(currentSlide) === "activity" ? `نشاط تفاعلي مباشر. المطلوب: ${getSlideMessages(currentSlide).join(" • ")}` : "") || "لا توجد ملاحظات محاضر خاصة لهذه الشريحة بعد."}
                    </p>
                  </div>
                  
                  <div className="grid grid-cols-2 gap-3 text-[10px] text-slate-500 mt-2">
                    <div className="flex items-center gap-1.5">
                      <Icons.CheckCircle className="w-3.5 h-3.5 text-emerald-500" />
                      <span>نوع المحتوى: <strong className="text-slate-800">{currentSlide ? getSlideType(currentSlide) : "content"}</strong></span>
                    </div>
                    <div className="flex items-center gap-1.5">
                      <Icons.Activity className="w-3.5 h-3.5 text-blue-500" />
                      <span>تأثير التفاعل: <strong className="text-slate-800">عالي ومبسط لجمهور الهواتف</strong></span>
                    </div>
                  </div>
                </div>
              )}
            </div>

            {/* Logo details and quick settings */}
            <div className="w-full md:w-64 bg-slate-50 p-4 rounded-2xl border border-slate-200 flex flex-col gap-3">
              <h4 className="text-xs font-bold text-slate-800 flex items-center gap-1.5 pb-2 border-b border-slate-200">
                <Icons.Settings className="w-4 h-4 text-blue-600" />
                خيارات العرض الذكي:
              </h4>

              <div>
                <label className="text-[10px] text-slate-500 mb-1.5 block">موقع شعار GLOBAL:</label>
                <div className="grid grid-cols-2 gap-1.5">
                  {[
                    { id: "bottom-right", label: "أسفل اليمين" },
                    { id: "top-left", label: "أعلى اليسار" }
                  ].map((pos) => (
                    <button
                      key={pos.id}
                      onClick={() => setLogoPosition(pos.id as any)}
                      className={`text-[9px] py-1 rounded-lg border transition-all ${logoPosition === pos.id ? "bg-blue-50 text-blue-600 border-blue-200 font-bold" : "bg-white text-slate-500 hover:text-slate-800 border-slate-100 shadow-sm"}`}
                    >
                      {pos.label}
                    </button>
                  ))}
                </div>
              </div>

              <div className="pt-2 border-t border-slate-200 flex items-center justify-between text-[10px] text-slate-400">
                <span>الحالة: متصل ومنظم</span>
                <span className="text-emerald-500 font-bold">● مباشر</span>
              </div>
            </div>

          </div>

        </section>

      </main>

      {/* 4. Complete Print Slide container for High-Quality PDF export (print-only) */}
      <div className="hidden print-container bg-white text-black p-0 m-0">
        {slides.map((slide) => (
          <div key={slide.id} className="print-slide-page flex flex-col justify-between" dir="rtl">
            
            {/* Top Bar for PDF */}
            <div className="flex items-start justify-between border-b border-slate-100 pb-4">
              <div className="flex items-center gap-4">
                <div className="text-amber-600 bg-amber-50 p-2.5 rounded-xl border border-amber-500/20">
                  {renderDynamicIcon(getSlideIcon(slide), "w-10 h-10")}
                </div>
                <div>
                  <span className="text-xs font-black text-amber-600 tracking-wider">
                    {getSlideLecture(slide)} • {getSlideSection(slide)}
                  </span>
                  <h3 className="text-sm font-bold text-slate-500 mt-1">
                    {slide.subtitle || "دليل المحاضر التفاعلي لعام 2026"}
                  </h3>
                </div>
              </div>
              
              <div className="text-right">
                <span className="text-xs font-black text-slate-400 font-mono block">الشريحة #{slide.id}</span>
                <span className="text-[10px] text-slate-400">GLOBAL FOR TRAINING</span>
              </div>
            </div>

            {/* Core Slide contents for PDF */}
            <div className="my-auto py-10 flex flex-col gap-6 justify-center">
              <h2 className="text-3xl font-black font-display text-slate-950 mb-4 leading-snug">
                {slide.title}
              </h2>
              
              <div className="flex flex-col gap-4">
                {getSlideMessages(slide).map((msg, index) => (
                  <div key={index} className="flex items-start gap-4">
                    <span className="text-amber-500 mt-1.5 text-lg shrink-0">✦</span>
                    <p className="text-lg font-bold text-slate-800 leading-relaxed">
                      {msg}
                    </p>
                  </div>
                ))}
              </div>
            </div>

            {/* Bottom Bar for PDF */}
            <div className="flex items-center justify-between pt-6 border-t border-slate-100">
              <span className="text-xs text-slate-400">
                حقوق الطبع محفوظة ورش عمل الشباب 2026 ©
              </span>
              
              <div className="flex items-center gap-2 bg-slate-950 text-white px-4 py-2 rounded-full">
                <span className="text-xs tracking-wider text-amber-400 font-black font-mono">GLOBAL FOR TRAINING</span>
              </div>
            </div>

          </div>
        ))}
      </div>

      {/* 5. Beautiful Interactive PDF Guide Modal (no-print) */}
      {showPrintIframeModal && (
        <div className="no-print fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-50 flex items-center justify-center p-4 transition-all">
          <div className="bg-white rounded-3xl max-w-lg w-full p-6 shadow-2xl border border-slate-100 flex flex-col gap-5 text-right animate-in fade-in zoom-in-95 duration-200" dir="rtl">
            
            {/* Modal Header */}
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2.5">
                <div className="w-10 h-10 rounded-full bg-blue-50 flex items-center justify-center text-blue-600">
                  <Icons.Printer className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-base font-extrabold text-slate-800">تعليمات تصدير الحقيبة التدريبية كـ PDF</h3>
                  <p className="text-xs text-slate-400 mt-0.5">خطوات بسيطة للحصول على أفضل دقة للطباعة</p>
                </div>
              </div>
              <button 
                onClick={() => setShowPrintIframeModal(false)}
                className="p-1.5 rounded-full hover:bg-slate-100 text-slate-400 hover:text-slate-600 transition-colors"
              >
                <Icons.X className="w-4 h-4" />
              </button>
            </div>

            {/* Modal Content */}
            <div className="flex flex-col gap-4 text-xs text-slate-600 leading-relaxed font-sans">
              
              <div className="p-3 bg-blue-50/50 rounded-2xl border border-blue-100 text-blue-800 font-medium">
                ⚠️ نظرًا لقيود الأمان وحماية الخصوصية التي تفرضها متصفحات الويب على النوافذ المضمنة (iFrame)، يرجى اتباع التعليمات التالية للتصدير بنجاح:
              </div>

              <div className="my-1">
                <button
                  onClick={() => {
                    window.open(window.location.href, '_blank');
                  }}
                  className="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs py-3 rounded-2xl flex items-center justify-center gap-2 transition-all shadow-md shadow-blue-600/10"
                >
                  <Icons.ExternalLink className="w-4 h-4" />
                  اضغط هنا لفتح الحقيبة في صفحة كاملة الآن
                </button>
              </div>

              <div className="space-y-3 mt-1">
                <div className="flex items-start gap-3">
                  <span className="w-5 h-5 rounded-full bg-blue-100 text-blue-600 font-bold flex items-center justify-center shrink-0 mt-0.5 text-[10px]">1</span>
                  <div>
                    <strong className="text-slate-800 block text-xs">اضغط على الزر الأزرق أعلاه:</strong>
                    <span>سيقوم بفتح الحقيبة مباشرة في علامة تبويب جديدة مستقلة بمتصفحك.</span>
                  </div>
                </div>

                <div className="flex items-start gap-3">
                  <span className="w-5 h-5 rounded-full bg-blue-100 text-blue-600 font-bold flex items-center justify-center shrink-0 mt-0.5 text-[10px]">2</span>
                  <div>
                    <strong className="text-slate-800 block text-xs">بدء عملية تصدير الـ PDF:</strong>
                    <span>في الصفحة الجديدة، اضغط على زر **"تصدير PDF عالي الجودة"** لتشغيل نافذة الطباعة التفاعلية.</span>
                  </div>
                </div>

                <div className="flex items-start gap-3">
                  <span className="w-5 h-5 rounded-full bg-blue-100 text-blue-600 font-bold flex items-center justify-center shrink-0 mt-0.5 text-[10px]">3</span>
                  <div>
                    <strong className="text-slate-800 block text-xs">تنسيق وضبط إعدادات الحفظ:</strong>
                    <span>في شاشة خيارات الطباعة بالمتصفح:</span>
                    <ul className="list-disc list-inside mr-4 mt-1 space-y-1 text-slate-500">
                      <li>اختر الوجهة: **حفظ بتنسيق PDF** (Save as PDF).</li>
                      <li>اختر الاتجاه: **أفقي** (Landscape).</li>
                      <li>تأكد من تفعيل خيار **"رسومات الخلفية"** (Background graphics) لإظهار كامل الأنماط البصرية والألوان الفاخرة للشرائح.</li>
                    </ul>
                  </div>
                </div>
              </div>

            </div>

            {/* Modal Footer */}
            <div className="flex justify-end gap-2 pt-3 border-t border-slate-100">
              <button
                onClick={() => setShowPrintIframeModal(false)}
                className="bg-blue-600 hover:bg-blue-750 text-white font-bold text-xs px-5 py-2.5 rounded-full shadow-sm shadow-blue-600/10 transition-colors"
              >
                حسناً، فهمت الطريقة
              </button>
            </div>

          </div>
        </div>
      )}

    </div>
  );
}

