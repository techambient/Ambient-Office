import { useState, useRef, useEffect } from 'react';
import { 
  PresentationData, 
  Slide, 
  SlideElement 
} from '../../types/office';
import { createDefaultSlide } from '../../utils/fileConverters';
import { 
  Play, 
  Plus, 
  Copy, 
  Trash2, 
  MoveUp, 
  MoveDown, 
  Type, 
  Square, 
  Image as ImageIcon, 
  Palette, 
  Layout, 
  X, 
  ChevronLeft, 
  ChevronRight,
  Maximize2,
  Sparkles,
  Bold,
  Italic,
  AlignLeft,
  AlignCenter,
  AlignRight
} from 'lucide-react';

interface PresentationEditorProps {
  data: PresentationData;
  onChange: (newData: PresentationData) => void;
}

export const PresentationEditor = ({ data, onChange }: PresentationEditorProps) => {
  const activeSlideIndex = data.slides.findIndex((s) => s.id === data.activeSlideId);
  const activeSlide = data.slides[activeSlideIndex] || data.slides[0];

  const [selectedElementId, setSelectedElementId] = useState<string | null>(null);
  const [isPlayingShow, setIsPlayingShow] = useState<boolean>(false);
  const [showSlideIndex, setShowSlideIndex] = useState<number>(0);
  const [showBackgroundPicker, setShowBackgroundPicker] = useState<boolean>(false);

  const selectedElement = activeSlide.elements.find((el) => el.id === selectedElementId);

  // Keyboard navigation for Slide Show
  useEffect(() => {
    const handleKey = (e: KeyboardEvent) => {
      if (isPlayingShow) {
        if (e.key === 'ArrowRight' || e.key === 'Space' || e.key === 'PageDown') {
          e.preventDefault();
          setShowSlideIndex((idx) => Math.min(data.slides.length - 1, idx + 1));
        } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
          e.preventDefault();
          setShowSlideIndex((idx) => Math.max(0, idx - 1));
        } else if (e.key === 'Escape') {
          setIsPlayingShow(false);
        }
      }
    };
    window.addEventListener('keydown', handleKey);
    return () => window.removeEventListener('keydown', handleKey);
  }, [isPlayingShow, data.slides.length]);

  const updateActiveSlide = (updates: Partial<Slide>) => {
    const newSlides = data.slides.map((s) =>
      s.id === activeSlide.id ? { ...s, ...updates } : s
    );
    onChange({ ...data, slides: newSlides });
  };

  const updateElement = (elId: string, updates: Partial<SlideElement>) => {
    const newElements = activeSlide.elements.map((el) =>
      el.id === elId ? { ...el, ...updates } : el
    );
    updateActiveSlide({ elements: newElements });
  };

  const handleAddSlide = () => {
    const newSlide = createDefaultSlide(data.slides.length + 1);
    const newSlides = [...data.slides, newSlide];
    onChange({
      ...data,
      slides: newSlides,
      activeSlideId: newSlide.id,
    });
  };

  const handleDuplicateSlide = () => {
    const newId = `slide-${Date.now()}`;
    const duplicated: Slide = {
      ...activeSlide,
      id: newId,
      title: `${activeSlide.title} (Copy)`,
      elements: activeSlide.elements.map((el) => ({ ...el, id: `el-${Date.now()}-${Math.random()}` })),
    };
    const newSlides = [...data.slides];
    newSlides.splice(activeSlideIndex + 1, 0, duplicated);
    onChange({ ...data, slides: newSlides, activeSlideId: newId });
  };

  const handleDeleteSlide = () => {
    if (data.slides.length <= 1) return;
    const newSlides = data.slides.filter((s) => s.id !== activeSlide.id);
    const nextActive = newSlides[Math.max(0, activeSlideIndex - 1)];
    onChange({ ...data, slides: newSlides, activeSlideId: nextActive.id });
  };

  const handleMoveSlide = (direction: 'up' | 'down') => {
    const targetIdx = direction === 'up' ? activeSlideIndex - 1 : activeSlideIndex + 1;
    if (targetIdx < 0 || targetIdx >= data.slides.length) return;
    const newSlides = [...data.slides];
    const [moved] = newSlides.splice(activeSlideIndex, 1);
    newSlides.splice(targetIdx, 0, moved);
    onChange({ ...data, slides: newSlides });
  };

  const handleAddText = () => {
    const newEl: SlideElement = {
      id: `el-text-${Date.now()}`,
      type: 'text',
      x: 20,
      y: 30,
      width: 60,
      height: 20,
      content: 'Click here to edit text',
      fontSize: 20,
      color: '#1e293b',
    };
    updateActiveSlide({ elements: [...activeSlide.elements, newEl] });
    setSelectedElementId(newEl.id);
  };

  const handleAddShape = () => {
    const newEl: SlideElement = {
      id: `el-shape-${Date.now()}`,
      type: 'shape',
      x: 25,
      y: 35,
      width: 50,
      height: 30,
      content: '',
      backgroundColor: '#3b82f6',
      borderColor: '#1d4ed8',
      borderWidth: 2,
      borderRadius: 12,
    };
    updateActiveSlide({ elements: [...activeSlide.elements, newEl] });
    setSelectedElementId(newEl.id);
  };

  const handleAddImage = () => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'image/*';
    input.onchange = (e: any) => {
      const file = e.target?.files?.[0];
      if (file) {
        const reader = new FileReader();
        reader.onload = (loadEvt) => {
          const base64 = loadEvt.target?.result as string;
          const newEl: SlideElement = {
            id: `el-img-${Date.now()}`,
            type: 'image',
            x: 20,
            y: 20,
            width: 50,
            height: 50,
            content: base64,
          };
          updateActiveSlide({ elements: [...activeSlide.elements, newEl] });
          setSelectedElementId(newEl.id);
        };
        reader.readAsDataURL(file);
      }
    };
    input.click();
  };

  const handleDeleteElement = (elId: string) => {
    updateActiveSlide({ elements: activeSlide.elements.filter((e) => e.id !== elId) });
    setSelectedElementId(null);
  };

  return (
    <div id="ambient-presentation-editor" className="flex flex-col h-full bg-slate-100 dark:bg-slate-950 overflow-hidden select-none">
      {/* Ribbon Navigation Bar */}
      <div className="bg-white dark:bg-slate-900 border-b border-slate-200 dark:border-slate-800 px-4 py-2 flex items-center justify-between shadow-2xs z-10 flex-wrap gap-2">
        {/* Slide Show and Add Tools */}
        <div className="flex items-center gap-2">
          <button
            id="btn-start-slideshow"
            onClick={() => {
              setShowSlideIndex(activeSlideIndex);
              setIsPlayingShow(true);
            }}
            className="flex items-center gap-1.5 px-3.5 py-1.5 text-xs font-semibold text-white bg-amber-600 hover:bg-amber-700 rounded-lg shadow-sm transition-colors cursor-pointer"
            title="Start Full-Screen Slide Show (F5)"
          >
            <Play className="w-3.5 h-3.5 fill-current" />
            <span>Slide Show</span>
          </button>

          <div className="h-4 w-px bg-slate-200 dark:border-slate-700 mx-1" />

          {/* Add Slide */}
          <button
            id="btn-add-slide"
            onClick={handleAddSlide}
            className="flex items-center gap-1 px-3 py-1.5 text-xs font-medium text-slate-700 dark:text-slate-200 bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 rounded-lg transition-colors"
          >
            <Plus className="w-3.5 h-3.5" />
            <span>New Slide</span>
          </button>

          {/* Add Elements */}
          <button
            id="btn-add-text"
            onClick={handleAddText}
            className="flex items-center gap-1 px-3 py-1.5 text-xs font-medium text-slate-700 dark:text-slate-200 bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 rounded-lg transition-colors"
          >
            <Type className="w-3.5 h-3.5 text-blue-600" />
            <span>Text Box</span>
          </button>

          <button
            id="btn-add-shape"
            onClick={handleAddShape}
            className="flex items-center gap-1 px-3 py-1.5 text-xs font-medium text-slate-700 dark:text-slate-200 bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 rounded-lg transition-colors"
          >
            <Square className="w-3.5 h-3.5 text-emerald-600" />
            <span>Shape</span>
          </button>

          <button
            id="btn-add-img"
            onClick={handleAddImage}
            className="flex items-center gap-1 px-3 py-1.5 text-xs font-medium text-slate-700 dark:text-slate-200 bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 rounded-lg transition-colors"
          >
            <ImageIcon className="w-3.5 h-3.5 text-purple-600" />
            <span>Image</span>
          </button>
        </div>

        {/* Slide Background & Theme */}
        <div className="flex items-center gap-2">
          {/* Background color */}
          <div className="flex items-center gap-1 text-xs">
            <span className="text-slate-500">Slide BG:</span>
            <input
              type="color"
              value={activeSlide.background || '#ffffff'}
              onChange={(e) => updateActiveSlide({ background: e.target.value })}
              className="w-6 h-6 rounded cursor-pointer border-0 bg-transparent"
              title="Change Slide Background Color"
            />
          </div>

          {/* Selected Element Controls (if selected) */}
          {selectedElement && (
            <div className="flex items-center gap-1.5 border-l border-slate-200 dark:border-slate-700 pl-2">
              <span className="text-xs font-medium text-slate-500">Element:</span>
              <input
                type="number"
                min={10}
                max={72}
                value={selectedElement.fontSize || 18}
                onChange={(e) => updateElement(selectedElement.id, { fontSize: Number(e.target.value) })}
                className="w-14 px-1.5 py-0.5 text-xs bg-slate-100 dark:bg-slate-800 rounded border border-slate-200 dark:border-slate-700"
                title="Font Size"
              />
              <input
                type="color"
                value={selectedElement.color || '#1e293b'}
                onChange={(e) => updateElement(selectedElement.id, { color: e.target.value })}
                className="w-5 h-5 rounded cursor-pointer border-0 bg-transparent"
                title="Text Color"
              />
              {selectedElement.type === 'shape' && (
                <input
                  type="color"
                  value={selectedElement.backgroundColor || '#3b82f6'}
                  onChange={(e) => updateElement(selectedElement.id, { backgroundColor: e.target.value })}
                  className="w-5 h-5 rounded cursor-pointer border-0 bg-transparent"
                  title="Shape Fill Color"
                />
              )}
              <button
                onClick={() => handleDeleteElement(selectedElement.id)}
                className="p-1 rounded text-red-500 hover:bg-red-50 dark:hover:bg-red-950/40"
                title="Delete Selected Element"
              >
                <Trash2 className="w-3.5 h-3.5" />
              </button>
            </div>
          )}
        </div>
      </div>

      {/* Main Workspace (Thumbnails + Stage) */}
      <div className="flex-1 flex overflow-hidden">
        {/* Left Thumbnails Sidebar */}
        <div className="w-56 bg-slate-50 dark:bg-slate-900 border-r border-slate-200 dark:border-slate-800 p-3 overflow-y-auto space-y-3 shrink-0">
          <div className="flex items-center justify-between text-xs font-semibold text-slate-500 px-1">
            <span>SLIDES ({data.slides.length})</span>
            <button
              onClick={handleAddSlide}
              className="text-amber-600 hover:text-amber-700 font-bold"
              title="Add slide"
            >
              + Add
            </button>
          </div>

          <div className="space-y-2">
            {data.slides.map((slide, idx) => {
              const isActive = slide.id === activeSlide.id;
              return (
                <div
                  key={slide.id}
                  id={`slide-thumb-${slide.id}`}
                  onClick={() => onChange({ ...data, activeSlideId: slide.id })}
                  className={`group relative p-2 rounded-xl border-2 transition-all cursor-pointer flex flex-col gap-1.5 ${
                    isActive
                      ? 'border-amber-500 bg-amber-50/40 dark:bg-amber-950/20 shadow-xs'
                      : 'border-slate-200 dark:border-slate-800 hover:border-slate-300 dark:hover:border-slate-700 bg-white dark:bg-slate-850'
                  }`}
                >
                  {/* Thumbnail mini canvas preview */}
                  <div
                    style={{ backgroundColor: slide.background || '#ffffff' }}
                    className="aspect-video w-full rounded-md border border-slate-200 dark:border-slate-700 flex items-center justify-center p-2 overflow-hidden shadow-2xs relative"
                  >
                    <span className="text-[10px] text-slate-400 font-semibold text-center truncate px-1">
                      {slide.title}
                    </span>
                  </div>

                  {/* Slide Label and Controls */}
                  <div className="flex items-center justify-between text-[11px]">
                    <span className="font-semibold text-slate-600 dark:text-slate-300">
                      {idx + 1}. {slide.title.substring(0, 14)}
                    </span>

                    {/* Reorder and Action buttons on hover */}
                    <div className="flex items-center gap-0.5 opacity-0 group-hover:opacity-100 transition-opacity">
                      {idx > 0 && (
                        <button
                          onClick={(e) => { e.stopPropagation(); handleMoveSlide('up'); }}
                          className="p-0.5 text-slate-400 hover:text-slate-700"
                          title="Move up"
                        >
                          <MoveUp className="w-3 h-3" />
                        </button>
                      )}
                      {idx < data.slides.length - 1 && (
                        <button
                          onClick={(e) => { e.stopPropagation(); handleMoveSlide('down'); }}
                          className="p-0.5 text-slate-400 hover:text-slate-700"
                          title="Move down"
                        >
                          <MoveDown className="w-3 h-3" />
                        </button>
                      )}
                      {data.slides.length > 1 && (
                        <button
                          onClick={(e) => { e.stopPropagation(); handleDeleteSlide(); }}
                          className="p-0.5 text-red-400 hover:text-red-600"
                          title="Delete slide"
                        >
                          <Trash2 className="w-3 h-3" />
                        </button>
                      )}
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </div>

        {/* Center Presentation Stage */}
        <div 
          className="flex-1 p-6 sm:p-12 overflow-auto flex items-center justify-center bg-slate-200/60 dark:bg-slate-950"
          onClick={() => setSelectedElementId(null)}
        >
          {/* 16:9 Slide Canvas */}
          <div
            id="active-slide-canvas"
            style={{ backgroundColor: activeSlide.background || '#ffffff' }}
            className="w-full max-w-[960px] aspect-video rounded-xl shadow-2xl border border-slate-300/80 dark:border-slate-800 relative overflow-hidden transition-all select-none"
          >
            {activeSlide.elements.map((el) => {
              const isSelected = selectedElementId === el.id;

              return (
                <div
                  key={el.id}
                  id={`slide-element-${el.id}`}
                  onClick={(e) => {
                    e.stopPropagation();
                    setSelectedElementId(el.id);
                  }}
                  style={{
                    left: `${el.x}%`,
                    top: `${el.y}%`,
                    width: `${el.width}%`,
                    height: `${el.height}%`,
                    backgroundColor: el.backgroundColor,
                    borderColor: el.borderColor,
                    borderWidth: el.borderWidth ? `${el.borderWidth}px` : undefined,
                    borderRadius: el.borderRadius ? `${el.borderRadius}px` : undefined,
                    color: el.color,
                    fontSize: el.fontSize ? `${el.fontSize}px` : undefined,
                    fontWeight: el.fontWeight,
                    textAlign: el.align || 'left',
                  }}
                  className={`absolute p-2.5 transition-shadow cursor-move ${
                    isSelected ? 'ring-2 ring-blue-500 ring-offset-2' : 'hover:outline-dashed hover:outline-1 hover:outline-blue-300'
                  }`}
                >
                  {el.type === 'heading' || el.type === 'text' ? (
                    <textarea
                      value={el.content}
                      onChange={(e) => updateElement(el.id, { content: e.target.value })}
                      className="w-full h-full bg-transparent border-0 outline-none resize-none overflow-hidden font-sans"
                      style={{ color: 'inherit', fontSize: 'inherit', fontWeight: 'inherit', textAlign: 'inherit' }}
                    />
                  ) : el.type === 'badge' ? (
                    <div className="w-full h-full flex items-center justify-center font-bold text-xs">
                      {el.content}
                    </div>
                  ) : el.type === 'image' ? (
                    <img src={el.content} alt="slide visual" className="w-full h-full object-contain pointer-events-none rounded" />
                  ) : null}
                </div>
              );
            })}
          </div>
        </div>
      </div>

      {/* Full-Screen Slide Show Mode */}
      {isPlayingShow && (
        <div className="fixed inset-0 z-50 bg-black flex flex-col items-center justify-center animate-in fade-in duration-200">
          {/* Slide Show Stage */}
          <div
            style={{ backgroundColor: data.slides[showSlideIndex]?.background || '#ffffff' }}
            className="w-full max-w-[1200px] aspect-video relative rounded-lg shadow-2xl overflow-hidden p-8"
          >
            {data.slides[showSlideIndex]?.elements.map((el) => (
              <div
                key={el.id}
                style={{
                  left: `${el.x}%`,
                  top: `${el.y}%`,
                  width: `${el.width}%`,
                  height: `${el.height}%`,
                  backgroundColor: el.backgroundColor,
                  borderColor: el.borderColor,
                  borderWidth: el.borderWidth ? `${el.borderWidth}px` : undefined,
                  borderRadius: el.borderRadius ? `${el.borderRadius}px` : undefined,
                  color: el.color,
                  fontSize: el.fontSize ? `${el.fontSize * 1.2}px` : undefined,
                  fontWeight: el.fontWeight,
                  textAlign: el.align || 'left',
                }}
                className="absolute p-3 whitespace-pre-line font-sans"
              >
                {el.type === 'image' ? (
                  <img src={el.content} alt="" className="w-full h-full object-contain" />
                ) : (
                  <span>{el.content}</span>
                )}
              </div>
            ))}
          </div>

          {/* Slide Show Navigation Overlay (bottom) */}
          <div className="absolute bottom-6 flex items-center gap-4 bg-slate-900/80 backdrop-blur-md px-5 py-2.5 rounded-full text-white text-xs shadow-xl border border-white/10">
            <button
              onClick={() => setShowSlideIndex((idx) => Math.max(0, idx - 1))}
              disabled={showSlideIndex === 0}
              className="p-1 hover:text-amber-400 disabled:opacity-30"
              title="Previous Slide"
            >
              <ChevronLeft className="w-5 h-5" />
            </button>

            <span className="font-mono font-bold tracking-wider">
              {showSlideIndex + 1} / {data.slides.length}
            </span>

            <button
              onClick={() => setShowSlideIndex((idx) => Math.min(data.slides.length - 1, idx + 1))}
              disabled={showSlideIndex === data.slides.length - 1}
              className="p-1 hover:text-amber-400 disabled:opacity-30"
              title="Next Slide"
            >
              <ChevronRight className="w-5 h-5" />
            </button>

            <div className="h-4 w-px bg-white/20 mx-1" />

            <button
              onClick={() => setIsPlayingShow(false)}
              className="flex items-center gap-1 hover:text-red-400"
              title="Exit Slide Show (Esc)"
            >
              <X className="w-4 h-4" />
              <span>Exit</span>
            </button>
          </div>
        </div>
      )}
    </div>
  );
};
