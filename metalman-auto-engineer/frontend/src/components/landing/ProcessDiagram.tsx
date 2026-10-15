import { Box, FileSpreadsheet, FileText, Cpu, ChevronDown } from "lucide-react";

const inputs = [
  { icon: Box, label: "CAD Assembly" },
  { icon: FileSpreadsheet, label: "Feasibility Sheet" },
  { icon: FileText, label: "Drawing PDFs" },
];

const outputs = [
  "Bill of Materials",
  "Process Flow Diagram",
  "Fixture Master List",
  "Fitment Checksheet",
  "Process FMEA",
  "Control Plan",
  "Fixture PM Plan",
];

export const ProcessDiagram = () => (
  <div className="relative rounded-2xl border border-white/10 bg-gradient-to-br from-white/5 to-transparent p-6 backdrop-blur-sm">
    <div className="rounded-xl border border-white/10 bg-slate-deep/40 p-6">
      <div className="mb-4 flex items-center justify-between text-[10px] uppercase tracking-wider text-white/40">
        <span>Your inputs</span>
        <span className="flex items-center gap-1.5">
          <span className="h-1.5 w-1.5 animate-pulse-glow rounded-full bg-success" /> Live pipeline
        </span>
      </div>

      {/* Tier 1: inputs */}
      <div className="grid grid-cols-3 gap-2">
        {inputs.map((i) => (
          <div
            key={i.label}
            className="flex flex-col items-center gap-1.5 rounded-lg border border-white/10 bg-white/[0.03] px-2 py-3 text-center"
          >
            <i.icon className="h-4 w-4 text-white/70" />
            <span className="text-[10px] leading-tight text-white/60">{i.label}</span>
          </div>
        ))}
      </div>

      <ChevronDown className="mx-auto my-2 h-4 w-4 text-white/30" />

      {/* Tier 2: engine */}
      <div className="flex items-center justify-center gap-2 rounded-lg border border-accent/30 bg-accent/10 px-4 py-2.5">
        <Cpu className="h-4 w-4 text-accent" />
        <span className="text-xs font-semibold text-white">Metalman Automation Engine</span>
      </div>
      <p className="mt-2 text-center text-[10px] text-white/40">
        Reads geometry, tolerances &amp; drawings — drafts every document automatically
      </p>

      <ChevronDown className="mx-auto my-2 h-4 w-4 text-white/30" />

      {/* Tier 3: outputs */}
      <div className="grid grid-cols-2 gap-2">
        {outputs.map((label) => (
          <div
            key={label}
            className="rounded-lg border border-white/10 bg-white/[0.02] px-2 py-2 text-[10px] leading-tight text-white/60"
          >
            {label}
          </div>
        ))}
      </div>

      {/* Golden-path caption */}
      <div className="mt-5 flex items-center justify-between border-t border-white/10 pt-4 text-[10px] uppercase tracking-wider text-white/40">
        <span>Upload</span>
        <span>→</span>
        <span>Select</span>
        <span>→</span>
        <span>Generate</span>
        <span>→</span>
        <span>Collect</span>
      </div>
    </div>
  </div>
);
