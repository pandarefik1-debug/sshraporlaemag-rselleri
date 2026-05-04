"use client";

import React, { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import * as htmlToImage from "html-to-image";
import { jsPDF } from "jspdf";
import * as pdfjsLib from "pdfjs-dist";
import pptxgen from "pptxgenjs";

if (typeof window !== "undefined") {
  pdfjsLib.GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjsLib.version}/build/pdf.worker.min.mjs`;
}

import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  PieChart,
  Pie,
  Cell,
  Legend,
} from "recharts";
import { UploadCloud, AlertCircle, Package, Wrench, RotateCcw, Download } from "lucide-react";

// --- TİP TANIMLAMALARI ---
interface ExcelRow {
  "Modül adı"?: string;
  "SSH MAMÜL TANIM"?: string;
  "MALZEME KISA METNİ"?: string;
  [key: string]: any;
}

interface ProcessedData {
  moduleName: string;
  productDesc: string;
  partDesc: string;
}

interface FrequencyData {
  name: string;
  count: number;
}

interface CombinationData {
  module: string;
  part: string;
  count: number;
}

// Recharts Grafik Renk Paleti (Kırmızı tonları)
const COLORS = ["#ef4444", "#dc2626", "#b91c1c", "#991b1b", "#7f1d1d"];

export default function QualityDashboard() {
  const [activeTab, setActiveTab] = useState<"dashboard" | "presentation">("dashboard");
  const [isLoaded, setIsLoaded] = useState<boolean>(false);
  const [isDragging, setIsDragging] = useState<boolean>(false);
  const [reportTitle, setReportTitle] = useState<string>("Fabrika Kalite Dashboard");
  const dashboardRef = useRef<HTMLDivElement>(null);
  
  // Sunum State'leri
  const [presentationFiles, setPresentationFiles] = useState<File[]>([]);
  const [isGeneratingPptx, setIsGeneratingPptx] = useState<boolean>(false);

  // Analiz Sonuçları State'leri
  const [totalRecords, setTotalRecords] = useState<number>(0);
  const [topModules, setTopModules] = useState<FrequencyData[]>([]);
  const [topParts, setTopParts] = useState<FrequencyData[]>([]);
  const [topCombos, setTopCombos] = useState<CombinationData[]>([]);

  // --- CORE LOGIC: EXCEL OKUMA VE HESAPLAMA ---
  const processExcel = (buffer: ArrayBuffer) => {
    const workbook = XLSX.read(buffer, { type: "array" });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    
    // JSON'a çevir, boş satırları atla
    const rawData = XLSX.utils.sheet_to_json<any>(worksheet, { defval: "" });

    const filtered: ProcessedData[] = [];
    
    // Sütun başlıklarındaki boşluk, büyük/küçük harf gibi tutarsızlıkları tolere etmek için
    const normalizeKey = (key: string) => key.toString().trim().toLowerCase().replace(/\s+/g, " ");
    
    const targetModuleKey = normalizeKey("Modül adı");
    const targetProductKey = normalizeKey("SSH MAMÜL TANIM");
    const targetPartKey = normalizeKey("MALZEME KISA METNİ");

    rawData.forEach((row) => {
      let moduleName = "";
      let productDesc = "";
      let partDesc = "";

      // Her satırdaki anahtarları normalize ederek kontrol et
      for (const [key, value] of Object.entries(row)) {
        const normalizedKey = normalizeKey(key);
        if (normalizedKey === targetModuleKey || normalizedKey.includes("modül ad")) {
          moduleName = String(value);
        } else if (normalizedKey === targetProductKey || normalizedKey.includes("mamül tan")) {
          productDesc = String(value);
        } else if (normalizedKey === targetPartKey || normalizedKey.includes("malzeme kısa")) {
          partDesc = String(value);
        }
      }

      if (moduleName.trim() && productDesc.trim() && partDesc.trim()) {
        filtered.push({
          moduleName: moduleName.trim(),
          productDesc: productDesc.trim(),
          partDesc: partDesc.trim(),
        });
      }
    });

    if (filtered.length === 0) {
      alert("Geçerli kayıt bulunamadı. Lütfen 'Modül adı', 'SSH MAMÜL TANIM' ve 'MALZEME KISA METNİ' sütunlarının varlığından emin olun.");
      return; // İşlemi durdur
    }

    calculateMetrics(filtered);
  };

  const calculateMetrics = (data: ProcessedData[]) => {
    setTotalRecords(data.length);

    const moduleCounts: Record<string, number> = {};
    const partCounts: Record<string, number> = {};
    const comboCounts: Record<string, number> = {};

    // Frekansları Hesapla
    data.forEach((item) => {
      moduleCounts[item.moduleName] = (moduleCounts[item.moduleName] || 0) + 1;
      partCounts[item.partDesc] = (partCounts[item.partDesc] || 0) + 1;

      // Çapraz Kombinasyon için anahtar oluştur
      const comboKey = `${item.moduleName} | ${item.partDesc}`;
      comboCounts[comboKey] = (comboCounts[comboKey] || 0) + 1;
    });

    // En Sorunlu 5 Modül
    const sortedModules = Object.entries(moduleCounts)
      .map(([name, count]) => ({ name, count }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 5);

    // En Sık Arızalanan 5 Parça
    const sortedParts = Object.entries(partCounts)
      .map(([name, count]) => ({ name, count }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 5);

    // En Yüksek Frekanslı 5 Kombinasyon
    const sortedCombos = Object.entries(comboCounts)
      .map(([key, count]) => {
        const [module, part] = key.split(" | ");
        return { module, part, count };
      })
      .sort((a, b) => b.count - a.count)
      .slice(0, 5);

    setTopModules(sortedModules);
    setTopParts(sortedParts);
    setTopCombos(sortedCombos);
    setIsLoaded(true);
  };

  // --- DRAG & DROP VE DOSYA YÜKLEME ---
  const handleDrop = useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files[0];
    if (file && (file.name.endsWith(".xlsx") || file.name.endsWith(".xls"))) {
      readFile(file);
    } else {
      alert("Lütfen geçerli bir Excel dosyası (.xlsx veya .xls) yükleyin.");
    }
  }, []);

  const handleFileUpload = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) readFile(file);
  }, []);

  const readFile = (file: File) => {
    const reader = new FileReader();
    reader.onload = (evt) => {
      if (evt.target?.result) {
        processExcel(evt.target.result as ArrayBuffer);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const resetDashboard = () => {
    setIsLoaded(false);
    setTotalRecords(0);
    setTopModules([]);
    setTopParts([]);
    setTopCombos([]);
  };

  const exportToPDF = async () => {
    if (!dashboardRef.current) return;
    
    try {
      const el = dashboardRef.current;
      
      // Elemanın ekrandaki GERÇEK genişliği ve tüm yüksekliğini (scroll dahil) alıyoruz.
      // Bu sayede sağda boşluk kalması veya içeriklerin kayması tamamen engellenir.
      const targetWidth = el.offsetWidth;
      const targetHeight = el.scrollHeight;
      
      const dataUrl = await htmlToImage.toPng(el, {
        quality: 1,
        backgroundColor: "#fef2f2",
        pixelRatio: 2,
        width: targetWidth, 
        height: targetHeight,
        style: {
          margin: "0",
          transform: "none",
          borderRadius: "0" // Tam kağıt hissi için oval kenarları kaldırıyoruz
        }
      });
      
      // Dikey (Portrait) A4 oluşturuyoruz
      const pdf = new jsPDF("p", "mm", "a4");
      const pageWidth = pdf.internal.pageSize.getWidth();
      const pageHeight = pdf.internal.pageSize.getHeight();
      
      // 1. ADIM: A4 Kağıdının TAMAMINI sitenin arka plan rengine (red-50) boyuyoruz.
      // Böylece kağıdın neresinde boşluk kalırsa kalsın sayfa bir bütün görünecek.
      pdf.setFillColor(254, 242, 242); // #fef2f2
      pdf.rect(0, 0, pageWidth, pageHeight, "F");
      
      const imgProps = pdf.getImageProperties(dataUrl);
      
      // 2. ADIM: Resmi sıfıra sıfır A4'e sığdıracak şekilde hesaplıyoruz
      let imgWidth = pageWidth;
      let imgHeight = (imgProps.height * imgWidth) / imgProps.width;
      
      // Eğer sayfa yüksekliğini aşıyorsa yüksekliğe göre sığdır
      if (imgHeight > pageHeight) {
        imgHeight = pageHeight;
        imgWidth = (imgProps.width * imgHeight) / imgProps.height;
      }
      
      // 3. ADIM: Kağıdın tam merkezine resmi yerleştiriyoruz
      const x = (pageWidth - imgWidth) / 2;
      const y = (pageHeight - imgHeight) / 2;
      
      pdf.addImage(dataUrl, "PNG", x, y, imgWidth, imgHeight);
      pdf.save("kalite-raporu.pdf");
    } catch (error: any) {
      console.error("PDF oluşturulurken hata:", error);
      alert("PDF oluşturulurken bir hata oluştu: " + error.message);
    }
  };

  // --- SUNUM BİRLEŞTİRİCİ LOGIC ---
  const handlePdfDrop = useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    const files = Array.from(e.dataTransfer.files).filter(f => f.name.toLowerCase().endsWith(".pdf"));
    if (files.length > 0) {
      setPresentationFiles(prev => [...prev, ...files]);
    } else {
      alert("Lütfen sadece PDF dosyaları yükleyin.");
    }
  }, []);

  const handlePdfUpload = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []).filter(f => f.name.toLowerCase().endsWith(".pdf"));
    if (files.length > 0) {
      setPresentationFiles(prev => [...prev, ...files]);
    }
  }, []);

  const generatePresentation = async () => {
    if (presentationFiles.length === 0) return;
    setIsGeneratingPptx(true);
    
    try {
      const pptx = new pptxgen();
      let isLayoutSet = false;
      let slideWidth = 10;
      let slideHeight = 5.625;
      
      for (const file of presentationFiles) {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument(arrayBuffer).promise;
        const numPages = pdf.numPages;
        
        for (let i = 1; i <= numPages; i++) {
          const page = await pdf.getPage(i);
          
          // İlk sayfadan orijinal PDF'in en-boy oranını alıp Sunum boyutunu ona göre ayarlıyoruz
          // Böylece grafiklerde sündürme/ezilme (ovalleşme) olmaz.
          if (!isLayoutSet) {
            const baseViewport = page.getViewport({ scale: 1 });
            // PDF point birimini inch'e çeviriyoruz (1 inch = 72 points)
            slideWidth = baseViewport.width / 72;
            slideHeight = baseViewport.height / 72;
            
            pptx.defineLayout({ name: 'CUSTOM_PDF_SIZE', width: slideWidth, height: slideHeight });
            pptx.layout = 'CUSTOM_PDF_SIZE';
            isLayoutSet = true;
          }
          
          // Görüntü Kalitesini Maksimuma Çıkarmak İçin Scale: 3 (Neredeyse 4K Kalitesi)
          const viewport = page.getViewport({ scale: 3 }); 
          
          const canvas = document.createElement("canvas");
          const context = canvas.getContext("2d");
          canvas.width = viewport.width;
          canvas.height = viewport.height;
          
          if (context) {
            // PDF'in şeffaf arkaplanı varsa JPEG'de siyaha dönmemesi için önce beyaza boyuyoruz
            context.fillStyle = "#ffffff";
            context.fillRect(0, 0, canvas.width, canvas.height);
            
            await page.render({ canvasContext: context, viewport: viewport } as any).promise;
            
            // Yüksek kalite JPEG olarak kaydediyoruz (PNG yaparsak çok ağır olur)
            const imgData = canvas.toDataURL("image/jpeg", 0.95);
            
            const slide = pptx.addSlide();
            // Slayda resmi tam oturtuyoruz, slayt zaten resimle aynı en-boy oranında!
            slide.addImage({ data: imgData, x: 0, y: 0, w: slideWidth, h: slideHeight });
          }
        }
      }
      
      await pptx.writeFile({ fileName: "Kalite_Toplu_Sunum.pptx" });
    } catch (error) {
      console.error("Sunum oluşturulurken hata:", error);
      alert("Sunum oluşturulurken bir hata oluştu. PDF'ler şifreli veya bozuk olabilir.");
    } finally {
      setIsGeneratingPptx(false);
    }
  };

  // --- RENDER ---
  return (
    <div className="min-h-screen bg-red-50 text-slate-900 p-6 md:p-10 font-sans">
      <div className="max-w-7xl mx-auto mb-8">
        {/* TABS (SEKMELER) */}
        <div className="flex space-x-2 bg-red-100/50 p-2 rounded-2xl mb-6 border border-red-200">
          <button
            onClick={() => setActiveTab("dashboard")}
            className={`flex-1 py-3 px-6 rounded-xl font-bold text-lg transition-all ${activeTab === "dashboard" ? "bg-white text-red-700 shadow-md border border-red-200" : "text-red-900/60 hover:bg-red-50 hover:text-red-700"}`}
          >
            📊 Kalite Dashboard
          </button>
          <button
            onClick={() => setActiveTab("presentation")}
            className={`flex-1 py-3 px-6 rounded-xl font-bold text-lg transition-all ${activeTab === "presentation" ? "bg-white text-red-700 shadow-md border border-red-200" : "text-red-900/60 hover:bg-red-50 hover:text-red-700"}`}
          >
            📑 Sunum Birleştirici
          </button>
        </div>

        {activeTab === "presentation" ? (
          /* SUNUM BİRLEŞTİRİCİ SEKMESİ */
          <div className="bg-white rounded-3xl p-10 border border-red-100 shadow-xl shadow-red-100/50 animate-in fade-in slide-in-from-bottom-4 duration-500">
            <div className="text-center mb-10">
              <h2 className="text-4xl font-black text-slate-900 mb-4">PDF to PowerPoint Birleştirici</h2>
              <p className="text-lg text-slate-500 max-w-2xl mx-auto">
                Analiz raporlarınızı ve diğer PDF dosyalarınızı buraya sürükleyin. Sistem tüm sayfaları birleştirip tek bir Sunum (.pptx) dosyası oluşturacaktır.
              </p>
            </div>

            <div
              onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
              onDragLeave={() => setIsDragging(false)}
              onDrop={handlePdfDrop}
              className={`border-2 border-dashed rounded-2xl p-16 text-center transition-all cursor-pointer flex flex-col items-center justify-center space-y-6 bg-red-50/30
                ${isDragging ? "border-red-500 bg-red-100" : "border-red-200 hover:border-red-400 hover:bg-red-50"}`}
            >
              <UploadCloud size={60} className={isDragging ? "text-red-500" : "text-slate-300"} />
              <div>
                <h3 className="text-2xl font-bold mb-2 text-slate-800">PDF Dosyalarını Sürükleyin</h3>
                <p className="text-slate-500">Aynı anda birden fazla PDF ekleyebilirsiniz.</p>
              </div>
              <input type="file" accept=".pdf" multiple className="hidden" id="pdf-upload" onChange={handlePdfUpload} />
              <label htmlFor="pdf-upload" className="px-6 py-3 bg-white border border-red-200 hover:border-red-400 text-red-600 rounded-xl font-bold cursor-pointer transition-colors shadow-sm">
                Dosya Seç
              </label>
            </div>

            {presentationFiles.length > 0 && (
              <div className="mt-10">
                <div className="flex justify-between items-center mb-6">
                  <h3 className="text-2xl font-bold text-slate-800">Seçilen Dosyalar ({presentationFiles.length})</h3>
                  <button 
                    onClick={() => setPresentationFiles([])}
                    className="text-sm font-bold text-slate-400 hover:text-red-500 transition-colors"
                  >
                    Tümünü Temizle
                  </button>
                </div>
                <ul className="space-y-3 mb-10">
                  {presentationFiles.map((f, i) => (
                    <li key={i} className="flex justify-between items-center bg-white p-4 rounded-xl border border-slate-100 shadow-sm">
                      <span className="font-bold text-slate-700 flex items-center gap-3">
                        <span className="bg-red-100 text-red-600 p-2 rounded-lg text-xs font-black">{i + 1}</span>
                        {f.name}
                      </span>
                      <button onClick={() => setPresentationFiles(prev => prev.filter((_, idx) => idx !== i))} className="text-red-400 hover:text-red-600 font-bold bg-red-50 px-3 py-1 rounded-lg">Kaldır</button>
                    </li>
                  ))}
                </ul>
                
                <div className="flex justify-center">
                  <button
                    onClick={generatePresentation}
                    disabled={isGeneratingPptx}
                    className="flex items-center space-x-3 bg-red-600 hover:bg-red-700 disabled:bg-red-300 text-white px-10 py-5 rounded-2xl transition-all shadow-xl font-black text-xl shadow-red-900/20"
                  >
                    <Download size={28} />
                    <span>{isGeneratingPptx ? "Sunum Oluşturuluyor Lütfen Bekleyin..." : "PowerPoint (.pptx) Olarak İndir"}</span>
                  </button>
                </div>
              </div>
            )}
          </div>
        ) : (
          /* KALİTE DASHBOARD SEKMESİ */
          <>
            {/* KONTROL PANOSU (PDF'te Çıkmaz, sadece web arayüzünde görünür) */}
            <div className="bg-white p-6 rounded-2xl border border-red-200 shadow-sm flex flex-col md:flex-row justify-between items-center gap-6 mb-8">
              <div className="flex-1 w-full">
                <label className="block text-xs font-bold text-red-500 uppercase tracking-wider mb-2">
                  Rapor Başlığı (Panodan Düzenle)
                </label>
                <input 
                  type="text"
                  value={reportTitle}
                  onChange={(e) => setReportTitle(e.target.value)}
                  className="w-full bg-red-50 border border-red-200 rounded-xl px-4 py-3 text-slate-900 font-bold focus:outline-none focus:ring-2 focus:ring-red-400 focus:border-red-400 transition-all text-lg placeholder-red-300"
                  placeholder="Örn: 2026 Mayıs Ayı Fabrika Kalite Raporu"
                />
              </div>
              
              {isLoaded && (
                <div className="flex items-center space-x-4 shrink-0">
                  <button
                    onClick={exportToPDF}
                    className="flex items-center space-x-2 bg-red-600 hover:bg-red-700 text-white px-6 py-3 rounded-xl transition-all shadow-md font-bold shadow-red-900/20"
                  >
                    <Download size={20} />
                    <span>PDF Olarak Çıktı Al</span>
                  </button>
                  <button
                    onClick={resetDashboard}
                    className="flex items-center space-x-2 bg-white hover:bg-red-50 text-slate-700 px-6 py-3 rounded-xl border border-red-200 transition-all shadow-sm font-bold"
                  >
                    <RotateCcw size={20} />
                    <span>Yeni Veri</span>
                  </button>
                </div>
              )}
            </div>

            {/* PDF İÇERİK ALANI KAPSAYICISI */}
            <div className="max-w-7xl mx-auto">
              {/* ASIL YAKALANACAK ALAN (A4 Tam sayfa efekti için yuvarlak köşe kaldırıldı) */}
              <div className="w-full bg-red-50 p-10 md:p-14" ref={dashboardRef}>
                
                {/* HEADER (PDF'in En Üstünde Görünecek Başlık) */}
                <div className="text-center mb-10 border-b-2 border-red-200 pb-8">
                  <h1 className="text-5xl font-black tracking-tight mb-4 text-red-900">{reportTitle || "Fabrika Kalite Dashboard"}</h1>
                  <p className="text-slate-500 text-xl font-medium">Satış Sonrası Hizmetler (SSH) Arıza Analizi</p>
                </div>

                {!isLoaded ? (
                  /* DOSYA YÜKLEME ALANI */
                  <div
                    onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
                    onDragLeave={() => setIsDragging(false)}
                    onDrop={handleDrop}
                    className={`border-2 border-dashed rounded-2xl p-20 text-center transition-all cursor-pointer flex flex-col items-center justify-center space-y-6 bg-white
                      ${isDragging ? "border-red-500 bg-red-50" : "border-red-200 hover:border-red-400 hover:bg-red-50"}`}
                  >
                    <UploadCloud size={80} className={isDragging ? "text-red-500" : "text-slate-300"} />
                    <div>
                      <h2 className="text-3xl font-bold mb-3 text-slate-800">Excel Verisini Sürükleyip Bırakın</h2>
                      <p className="text-slate-500 text-lg">Sadece .xlsx ve .xls dosyaları desteklenmektedir.</p>
                    </div>
                    <input
                      type="file"
                      accept=".xlsx, .xls"
                      className="hidden"
                      id="file-upload"
                      onChange={handleFileUpload}
                    />
                    <label
                      htmlFor="file-upload"
                      className="px-8 py-4 bg-red-600 hover:bg-red-700 rounded-xl font-bold text-lg cursor-pointer transition-colors shadow-lg shadow-red-900/20 mt-4 inline-block"
                    >
                      Bilgisayardan Seç
                    </label>
                  </div>
                ) : (
                  /* ANALİZ SONUÇLARI KISMI */
                  <div className="space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-700">
                    
                    {/* KPI KARTLARI */}
                    <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                      {/* Kart 1: Toplam Kayıt */}
                      <div className="bg-white rounded-2xl p-8 border border-red-100 shadow-xl shadow-red-100/50 flex items-center space-x-6">
                        <div className="p-5 bg-red-50 rounded-xl text-red-600 border border-red-100">
                          <AlertCircle size={40} />
                        </div>
                        <div className="flex-1 min-w-0">
                          <p className="text-slate-500 text-sm font-bold uppercase tracking-widest mb-1">Toplam SSH Kaydı</p>
                          <p className="text-5xl font-black text-slate-900">{totalRecords.toLocaleString("tr-TR")}</p>
                        </div>
                      </div>

                      {/* Kart 2: En Sorunlu Modül */}
                      <div className="bg-white rounded-2xl p-8 border border-red-100 shadow-xl shadow-red-100/50 flex items-center space-x-6">
                        <div className="p-5 bg-red-50 rounded-xl text-red-600 border border-red-100">
                          <Package size={40} />
                        </div>
                        <div className="flex-1 min-w-0">
                          <p className="text-slate-500 text-sm font-bold uppercase tracking-widest mb-1">En Sorunlu Modül</p>
                          <p className="text-3xl font-bold text-slate-900 truncate" title={topModules[0]?.name}>
                            {topModules[0]?.name || "-"}
                          </p>
                        </div>
                      </div>

                      {/* Kart 3: En Sık Arızalanan Parça */}
                      <div className="bg-white rounded-2xl p-8 border border-red-100 shadow-xl shadow-red-100/50 flex items-center space-x-6">
                        <div className="p-5 bg-red-50 rounded-xl text-red-600 border border-red-100">
                          <Wrench size={40} />
                        </div>
                        <div className="flex-1 min-w-0">
                          <p className="text-slate-500 text-sm font-bold uppercase tracking-widest mb-1">En Sık Arızalanan Parça</p>
                          <p className="text-2xl font-bold text-slate-900 line-clamp-2 leading-tight" title={topParts[0]?.name}>
                            {topParts[0]?.name || "-"}
                          </p>
                        </div>
                      </div>
                    </div>

                    {/* GRAFİKLER */}
                    <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
                      
                      {/* Bar Chart (Modüller) */}
                      <div className="bg-white rounded-2xl p-8 border border-red-100 shadow-xl shadow-red-100/50">
                        <h3 className="text-2xl font-bold mb-8 text-slate-900">Modüllere Göre Hata Dağılımı (İlk 5)</h3>
                        <div className="h-[350px] w-full">
                          <ResponsiveContainer width="100%" height="100%">
                            <BarChart data={topModules} margin={{ top: 20, right: 30, left: -20, bottom: 5 }}>
                              <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" vertical={false} />
                              <XAxis dataKey="name" stroke="#64748b" tick={{ fill: "#64748b", fontSize: 13 }} tickLine={false} axisLine={false} />
                              <YAxis stroke="#64748b" tick={{ fill: "#64748b", fontSize: 13 }} tickLine={false} axisLine={false} />
                              <Tooltip
                                cursor={{ fill: "#f1f5f9", opacity: 0.4 }}
                                contentStyle={{ backgroundColor: "#ffffff", borderColor: "#fecaca", color: "#0f172a", borderRadius: "0.75rem", padding: "12px", fontWeight: "bold" }}
                                itemStyle={{ color: "#ef4444", fontWeight: "bold" }}
                              />
                              <Bar dataKey="count" fill="#ef4444" radius={[6, 6, 0, 0]} animationDuration={1500}>
                                {topModules.map((entry, index) => (
                                  <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                                ))}
                              </Bar>
                            </BarChart>
                          </ResponsiveContainer>
                        </div>
                      </div>

                      {/* Pie Chart (Parçalar) */}
                      <div className="bg-white rounded-2xl p-8 border border-red-100 shadow-xl shadow-red-100/50">
                        <h3 className="text-2xl font-bold mb-8 text-slate-900">Malzemelere Göre Hata Dağılımı (İlk 5)</h3>
                        <div className="h-[350px] w-full">
                          <ResponsiveContainer width="100%" height="100%">
                            <PieChart>
                              <Pie
                                data={topParts}
                                cx="50%"
                                cy="50%"
                                innerRadius={80}
                                outerRadius={130}
                                paddingAngle={5}
                                dataKey="count"
                                stroke="none"
                                animationDuration={1500}
                              >
                                {topParts.map((entry, index) => (
                                  <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                                ))}
                              </Pie>
                              <Tooltip
                                contentStyle={{ backgroundColor: "#ffffff", borderColor: "#fecaca", color: "#0f172a", borderRadius: "0.75rem", padding: "12px", fontWeight: "bold" }}
                              />
                              <Legend 
                                wrapperStyle={{ paddingTop: "20px" }} 
                                formatter={(value) => <span className="text-slate-700 font-medium">{value}</span>}
                              />
                            </PieChart>
                          </ResponsiveContainer>
                        </div>
                      </div>
                    </div>

                    {/* ÇAPRAZ KOMBİNASYON TABLOSU */}
                    <div className="bg-white rounded-2xl border border-red-100 shadow-xl shadow-red-100/50 overflow-hidden">
                      <div className="p-8 border-b border-red-100 bg-red-50/50">
                        <h3 className="text-2xl font-bold text-slate-900">Çapraz Analiz: En Çok Sorun Çıkaran Modül & Parça Eşleşmeleri</h3>
                        <p className="text-slate-500 mt-2">Hangi ürünün, hangi parçasında en çok sorun çıkıyor?</p>
                      </div>
                      <div className="overflow-x-auto">
                        <table className="w-full text-left border-collapse">
                          <thead className="bg-red-50/80">
                            <tr className="text-red-800 text-sm uppercase tracking-widest border-b border-red-100">
                              <th className="p-6 font-bold w-24">Sıra</th>
                              <th className="p-6 font-bold">Modül Adı</th>
                              <th className="p-6 font-bold">Parça Adı (Malzeme Kısa Metni)</th>
                              <th className="p-6 font-bold text-right">Hata Sayısı</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-red-50">
                            {topCombos.map((combo, index) => (
                              <tr key={index} className="hover:bg-red-50 transition-colors group">
                                <td className="p-6 font-bold text-slate-400 group-hover:text-red-500 transition-colors">#{index + 1}</td>
                                <td className="p-6 font-bold text-xl text-slate-800">{combo.module}</td>
                                <td className="p-6 text-lg text-slate-600">{combo.part}</td>
                                <td className="p-6 text-right">
                                  <span className="inline-flex items-center justify-center px-4 py-1.5 rounded-full bg-red-100 text-red-700 font-black text-2xl border border-red-200">
                                    {combo.count}
                                  </span>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>

                  </div>
                )}
              </div>
            </div>
          </>
        )}
      </div>
    </div>
  );
}
