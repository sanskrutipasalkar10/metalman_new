import { useState } from "react";
import { useNavigate } from "react-router-dom";
import { Bell, ChevronDown, FileSpreadsheet, FileText, LogOut, Sparkles, LayoutDashboard, Settings, History, Crop } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Checkbox } from "@/components/ui/checkbox";
import { Logo } from "@/components/Logo";
import { UploadZone } from "@/components/UploadZone";
import { ProcessingView } from "@/components/ProcessingView";
import { Avatar, AvatarFallback } from "@/components/ui/avatar";
import { cn } from "@/lib/utils";
import { API_BASE } from "@/lib/api";

const OUTPUTS = [
  { id: "bom",          label: "Bill of Materials (BOM)",        desc: "Consolidated list of parts for NPD costing and procurement" },
  { id: "pfd",          label: "Process Flow Diagram (PFD)", desc: "Sequential manufacturing workflow for PPAP approval" },
  { id: "tooling",      label: "Fixture Master List",            desc: "Complete inventory of jigs, fixtures, and shop floor tooling" },
  { id: "fitment",      label: "Fitment Checksheet",             desc: "Part fitment verification sheet with part images and BOM data" },
  { id: "pfmea",        label: "Process FMEA (All Sheets)",      desc: "Failure Mode & Effects Analysis for Assembly, SM & BOP" },
  { id: "control_plan", label: "Control Plan (All Sheets)",      desc: "Contractual Control Plan: Sub Assy Index, Assy, Sheetmetal & BOP" },
  { id: "fixture_plan", label: "Fixture PM Master Plan",          desc: "Monthly fixture preventive maintenance schedule with day-wise tracking (1–31)" },
];

const Dashboard = () => {
  const navigate = useNavigate();
  const [feasibilityFile, setFeasibilityFile] = useState<File | null>(null);
  const [cadFile, setCadFile] = useState<File | null>(null);
  const [drawingFiles, setDrawingFiles] = useState<File[]>([]);
  
  const [outputs, setOutputs] = useState<Record<string, boolean>>({ bom: true, pfd: true, tooling: false, fitment: false, pfmea: false, control_plan: false, fixture_plan: false });
  const [processing, setProcessing] = useState(false);
  const [taskId, setTaskId] = useState<string | null>(null);

  const canGenerate = feasibilityFile && cadFile && !processing;

  const handleGenerate = async () => {
    if (!feasibilityFile || !cadFile) return;
    
    setProcessing(true);
    const formData = new FormData();
    formData.append("feasibility_file", feasibilityFile);
    formData.append("cad_file", cadFile);
    
    drawingFiles.forEach((file) => {
      formData.append("drawings", file);
    });
    
    formData.append("outputs", JSON.stringify(outputs));
    
    try {
      const resp = await fetch(`${API_BASE}/api/analyze`, {
        method: "POST",
        body: formData,
      });
      const data = await resp.json();
      if (data.task_id) {
        setTaskId(data.task_id);
      } else {
        throw new Error(data.error || "Failed to start task");
      }
    } catch (err) {
      console.error("Generation failed:", err);
      setProcessing(false);
      alert("Failed to start engineering pipeline. Please check backend.");
    }
  };

  const reset = () => {
    setProcessing(false);
    setTaskId(null);
    setFeasibilityFile(null);
    setCadFile(null);
    setDrawingFiles([]);
  };

  return (
    <div className="min-h-screen bg-secondary/30">
      <nav className="sticky top-0 z-50 border-b border-border bg-card/80 backdrop-blur-md">
        <div className="container flex h-16 items-center justify-between">
          <div className="flex items-center gap-8">
            <Logo className="h-8" />
            <div className="hidden h-8 w-px bg-border md:block" />
            <nav className="hidden items-center gap-6 md:flex">
              <a href="/dashboard" className="flex items-center gap-2 text-sm font-medium text-primary">
                <LayoutDashboard className="h-4 w-4" /> Dashboard
              </a>
              <a href="/cropper" className="flex items-center gap-2 text-sm font-medium text-muted-foreground transition-colors hover:text-foreground">
                <Crop className="h-4 w-4" /> Cropper
              </a>

              <a href="/stitch" className="flex items-center gap-2 rounded-full bg-primary/10 px-3 py-1.5 text-sm font-bold text-primary transition-all hover:bg-primary/20 hover:shadow-glow">
                <Sparkles className="h-4 w-4" /> Application Demo
              </a>

              <a href="#" className="flex items-center gap-2 text-sm font-medium text-muted-foreground transition-colors hover:text-foreground">
                <History className="h-4 w-4" /> History
              </a>
              <a href="#" className="flex items-center gap-2 text-sm font-medium text-muted-foreground transition-colors hover:text-foreground">
                <Settings className="h-4 w-4" /> Settings
              </a>
            </nav>
          </div>
          <div className="flex items-center gap-4">
            <Button variant="ghost" size="icon" className="relative">
              <Bell className="h-5 w-5 text-muted-foreground" />
              <span className="absolute right-2 top-2 h-2 w-2 rounded-full bg-accent" />
            </Button>
            
            <Button 
              variant="ghost" 
              size="icon" 
              onClick={() => {
                localStorage.removeItem("metalman_token");
                localStorage.removeItem("metalman_user");
                window.location.href = "/login";
              }}
              title="Logout"
            >
              <LogOut className="h-5 w-5 text-muted-foreground" />
            </Button>

            <Avatar className="h-9 w-9 border border-border shadow-sm">
              <AvatarFallback className="bg-primary/5 text-xs font-bold">ME</AvatarFallback>
            </Avatar>
          </div>
        </div>
      </nav>

      <main className="container py-10">
        <div className="mx-auto max-w-3xl">
          {processing && taskId ? (
            <ProcessingView taskId={taskId} onReset={reset} />
          ) : (
            <div className="overflow-hidden rounded-2xl border border-border bg-card shadow-card animate-in fade-in slide-in-from-bottom-4 duration-500">
              <section className="border-b border-border p-7">
                <SectionHeader step="01" title="Input Files" subtitle="Upload your engineering source data" />
                <div className="mt-5 space-y-4">
                  <UploadZone
                    step={1}
                    title="NPD Feasibility Matrix"
                    accept=".xlsx"
                    hint="Upload the signed feasibility Excel (Tolerance, Material, Capacity)."
                    onChangeFile={(f) => setFeasibilityFile(f)}
                  />
                  <UploadZone
                    step={2}
                    title="Master assembly (.STP)"
                    accept=".stp,.step"
                    hint="The master 3D STEP file for geometric slicing."
                    onChangeFile={(f) => setCadFile(f)}
                  />
                  <UploadZone
                    step={3}
                    title="Technical Drawings / PDF"
                    accept=".pdf"
                    multiple={true}
                    hint="2D drawings for vision-based isometric extraction."
                    onChangeFiles={(fs) => setDrawingFiles(fs)}
                  />
                </div>
              </section>
              
              <section className="border-b border-border p-7">
                <SectionHeader step="02" title="Requested Outputs" subtitle="Select manufacturing documents to generate" />
                <div className="mt-5 space-y-3">
                  {OUTPUTS.map((output) => (
                    <div
                      key={output.id}
                      onClick={() => setOutputs(prev => ({ ...prev, [output.id]: !prev[output.id] }))}
                      className={cn(
                        "group flex cursor-pointer items-center gap-4 rounded-xl border p-4 transition-all duration-200",
                        outputs[output.id] 
                          ? "border-primary/50 bg-primary/5 shadow-sm" 
                          : "border-border hover:border-primary/30 hover:bg-secondary/40"
                      )}
                    >
                      <div className={cn(
                        "flex h-10 w-10 items-center justify-center rounded-lg border transition-colors",
                        outputs[output.id] ? "border-primary bg-primary text-primary-foreground" : "bg-secondary text-muted-foreground"
                      )}>
                        {output.id === 'bom' ? <FileSpreadsheet className="h-5 w-5" /> : <FileText className="h-5 w-5" />}
                      </div>
                      <div className="flex-1">
                        <div className="text-sm font-bold">{output.label}</div>
                        <div className="text-xs text-muted-foreground">{output.desc}</div>
                      </div>
                      <Checkbox 
                        checked={outputs[output.id]} 
                        onCheckedChange={() => {}} // Handled by div click
                        className="h-5 w-5 rounded-md border-2"
                      />
                    </div>
                  ))}
                </div>
              </section>

              <section className="bg-secondary/40 p-7">
                <Button
                  disabled={!canGenerate}
                  onClick={handleGenerate}
                  className={cn(
                    "h-14 w-full text-base font-bold transition-all duration-300", 
                    canGenerate 
                      ? "bg-accent text-accent-foreground shadow-orange hover:translate-y-[-2px] hover:shadow-orange-lg" 
                      : "bg-muted"
                  )}
                >
                  {processing ? (
                    <span className="flex items-center gap-2">
                      <Sparkles className="h-5 w-5 animate-spin" /> Starting NPD Pipeline...
                    </span>
                  ) : canGenerate ? (
                    "Launch Automated Generation"
                  ) : (
                    "Upload NPD Source Files"
                  )}
                </Button>
                <p className="mt-4 text-center text-[11px] text-muted-foreground uppercase tracking-widest font-semibold opacity-60">
                  Engineering Automation Framework • Metalman Auto Ltd
                </p>
              </section>
            </div>
          )}
        </div>
      </main>
    </div>
  );
};

const SectionHeader = ({ step, title, subtitle }: { step: string; title: string; subtitle: string }) => (
  <div className="flex items-start gap-4">
    <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-md bg-primary font-mono text-xs font-bold text-primary-foreground shadow-sm">
      {step}
    </div>
    <div>
      <h2 className="text-lg font-bold leading-tight">{title}</h2>
      <p className="mt-0.5 text-sm text-muted-foreground">{subtitle}</p>
    </div>
  </div>
);

export default Dashboard;