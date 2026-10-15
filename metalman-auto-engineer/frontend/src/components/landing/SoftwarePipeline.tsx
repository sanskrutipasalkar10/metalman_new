import {
  Box,
  FileText,
  FileSpreadsheet,
  Layers,
  Database,
  Scissors,
  ScanLine,
  FileOutput,
  ShieldAlert,
  ClipboardCheck,
  Award,
  Ruler,
  Sparkles,
  ShieldCheck,
  type LucideIcon,
} from "lucide-react";
import { RevealNode } from "@/components/landing/RevealNode";
import { HoverCard, HoverCardContent, HoverCardTrigger } from "@/components/ui/hover-card";
import { cn } from "@/lib/utils";

interface Bullet {
  icon: LucideIcon;
  text: string;
  standard?: string;
}

type Stage = {
  icons: LucideIcon[];
  label: string;
  sub: string;
} & (
  | { detailType: "bullets"; bullets: Bullet[] }
  | { detailType: "chips"; chips: string[]; standardsNote: string }
);

const stages: Stage[] = [
  {
    icons: [Box, FileText, FileSpreadsheet],
    label: "Input Files",
    sub: "CAD, drawings & master data",
    detailType: "bullets",
    bullets: [
      { icon: Box, text: "CAD Assembly (.STEP) — the 3D master model", standard: "ASME Y14.5" },
      { icon: FileSpreadsheet, text: "Feasibility Sheet — part numbers, routing & tolerances", standard: "AIAG APQP" },
      { icon: FileText, text: "Drawing PDFs — dimensional & isometric references", standard: "ASME Y14.5" },
    ],
  },
  {
    icons: [Layers],
    label: "Template Library",
    sub: "Saved document templates",
    detailType: "chips",
    chips: [
      "BOM Template",
      "PFD Template",
      "Tooling List Template",
      "Fitment Checksheet Template",
      "PFMEA Template",
      "Control Plan Template",
      "Fixture Plan Template",
    ],
    standardsNote: "Formatted to AIAG-VDA FMEA Handbook, AIAG APQP & Control Plan Manual, and ASME Y14.5 conventions",
  },
  {
    icons: [Database],
    label: "Knowledge Base",
    sub: "Engineering reference data",
    detailType: "bullets",
    bullets: [
      { icon: ShieldAlert, text: "Failure-mode libraries — Assembly, Sheet Metal & BOP processes", standard: "AIAG-VDA FMEA" },
      { icon: ClipboardCheck, text: "Pre-approved Control Plan checkpoints per process type", standard: "IATF 16949" },
      { icon: Award, text: "Years of Metalman engineering standards, reused on every part" },
    ],
  },
  {
    icons: [Scissors, ScanLine],
    label: "Processing",
    sub: "Cuts, crops & cross-checks",
    detailType: "bullets",
    bullets: [
      { icon: ScanLine, text: "Computer vision extracts & cleans part views from drawings" },
      { icon: Box, text: "CAD engine slices the assembly operation-by-operation" },
      { icon: Ruler, text: "Cross-checks tolerances & part numbers against your spec sheet", standard: "ASME Y14.5" },
      { icon: Sparkles, text: "AI steps in only when deterministic logic can't find a view" },
    ],
  },
  {
    icons: [FileOutput],
    label: "Generated Documents",
    sub: "Ready-to-submit package",
    detailType: "chips",
    chips: [
      "Bill of Materials",
      "Process Flow Diagram",
      "Tooling / Fixture List",
      "Fitment Checksheet",
      "Process FMEA",
      "Control Plan",
      "Fixture PM Plan",
    ],
    standardsNote: "Delivered per AIAG-VDA FMEA Handbook, AIAG APQP & Control Plan Manual, and ASME Y14.5 dimensioning standards",
  },
];

const dashedLineX = {
  backgroundImage:
    "repeating-linear-gradient(to right, hsl(var(--border)) 0, hsl(var(--border)) 8px, transparent 8px, transparent 16px)",
};
const dashedLineY = {
  backgroundImage:
    "repeating-linear-gradient(to bottom, hsl(var(--border)) 0, hsl(var(--border)) 8px, transparent 8px, transparent 16px)",
};

const StageNode = ({ stage }: { stage: Stage }) => (
  <button
    type="button"
    className="group relative flex w-full flex-col items-center gap-3 rounded-xl border border-border bg-card p-6 text-center shadow-card transition-all duration-300 hover:-translate-y-1 hover:border-primary/40 hover:shadow-steel focus-visible:-translate-y-1 focus-visible:border-primary/40 focus-visible:shadow-steel focus:outline-none md:w-44"
  >
    <span className="absolute right-3 top-3 h-1.5 w-1.5 animate-pulse-glow rounded-full bg-primary/60 group-hover:bg-accent" />
    <div className="flex h-14 w-14 items-center justify-center gap-1 rounded-lg bg-secondary text-primary transition-colors group-hover:bg-primary group-hover:text-primary-foreground">
      {stage.icons.map((Icon, i) => (
        <Icon key={i} className={stage.icons.length > 1 ? "h-4 w-4" : "h-6 w-6"} />
      ))}
    </div>
    <div>
      <div className="text-sm font-bold">{stage.label}</div>
      <div className="mt-1 text-xs text-muted-foreground">{stage.sub}</div>
    </div>
  </button>
);

const StageDetails = ({ stage, index }: { stage: Stage; index: number }) => {
  const HeaderIcon = stage.icons[0];
  return (
    <>
      <div className="h-1.5 bg-gradient-to-r from-primary via-primary-glow to-accent" />
      <div className="p-5">
        <div className="mb-4 flex items-center gap-3">
          <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-lg bg-gradient-to-br from-primary to-primary-glow text-white shadow-steel">
            <HeaderIcon className="h-5 w-5" />
          </div>
          <div>
            <div className="text-sm font-bold">{stage.label}</div>
            <div className="font-mono text-[10px] uppercase tracking-wider text-muted-foreground">
              Step {index + 1} of {stages.length}
            </div>
          </div>
        </div>

        {stage.detailType === "bullets" ? (
          <ul className="space-y-3">
            {stage.bullets.map((item) => (
              <li key={item.text} className="flex items-start gap-2.5">
                <span className="mt-0.5 flex h-6 w-6 shrink-0 items-center justify-center rounded-md bg-secondary text-primary">
                  <item.icon className="h-3.5 w-3.5" />
                </span>
                <span className="flex-1 text-sm leading-snug text-muted-foreground">
                  {item.text}
                  {item.standard && (
                    <span className="ml-1.5 inline-block whitespace-nowrap rounded border border-border bg-muted px-1.5 py-0.5 align-middle font-mono text-[9px] font-semibold uppercase tracking-wide text-muted-foreground/80">
                      {item.standard}
                    </span>
                  )}
                </span>
              </li>
            ))}
          </ul>
        ) : (
          <>
            <div className="flex flex-wrap gap-1.5">
              {stage.chips.map((item) => (
                <span
                  key={item}
                  className="rounded-full border border-primary/20 bg-secondary px-2.5 py-1 text-xs font-medium text-primary"
                >
                  {item}
                </span>
              ))}
            </div>
            <div className="mt-4 flex items-start gap-2 border-t border-border pt-3 text-[10px] leading-relaxed text-muted-foreground">
              <ShieldCheck className="mt-0.5 h-3.5 w-3.5 shrink-0 text-accent" />
              <span>{stage.standardsNote}</span>
            </div>
          </>
        )}
      </div>
    </>
  );
};

export const SoftwarePipeline = () => (
  <section id="features" className="border-b border-border bg-background py-24">
    <div className="container">
      <div className="mx-auto mb-16 max-w-2xl text-center">
        <div className="mb-4 text-xs font-semibold uppercase tracking-[0.2em] text-primary">How It Works</div>
        <h2 className="text-balance text-4xl font-bold md:text-5xl">From raw files to finished paperwork.</h2>
        <p className="mt-4 text-lg text-muted-foreground">
          No manual copying. The engine reads your files, matches them against your own templates and
          engineering standards, and hands back the finished package.
        </p>
        <p className="mt-2 text-xs font-medium uppercase tracking-wide text-muted-foreground/70">
          Hover any step for details
        </p>
      </div>

      <div className="mx-auto flex max-w-5xl flex-col items-stretch md:flex-row md:items-center">
        {stages.map((stage, i) => (
          <div key={stage.label} className="flex flex-col items-center md:flex-1 md:flex-row">
            <RevealNode index={i} className="w-full md:w-auto">
              <HoverCard openDelay={80} closeDelay={80}>
                <HoverCardTrigger asChild>
                  <div>
                    <StageNode stage={stage} />
                  </div>
                </HoverCardTrigger>
                <HoverCardContent
                  className={cn(
                    "w-72 overflow-hidden rounded-xl border-border/60 p-0 shadow-elevated sm:w-80",
                    stage.detailType === "chips" && "w-80 sm:w-96"
                  )}
                >
                  <StageDetails stage={stage} index={i} />
                </HoverCardContent>
              </HoverCard>
            </RevealNode>
            {i < stages.length - 1 && (
              <>
                <div className="my-2 h-8 w-px animate-flow-line-y md:hidden" style={dashedLineY} />
                <div className="hidden h-px flex-1 animate-flow-line-x md:mx-2 md:block" style={dashedLineX} />
              </>
            )}
          </div>
        ))}
      </div>
    </div>
  </section>
);
