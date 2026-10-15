import {
  Target,
  PenTool,
  FlaskConical,
  Factory,
  FileSpreadsheet,
  Route,
  AlertTriangle,
  ClipboardCheck,
  FolderOpen,
  CheckCircle2,
  Handshake,
  type LucideIcon,
} from "lucide-react";
import { Badge } from "@/components/ui/badge";
import { RevealNode } from "@/components/landing/RevealNode";
import { ApqpDeepDive } from "@/components/landing/ApqpDeepDive";
import { PpapDeepDive } from "@/components/landing/PpapDeepDive";

interface Step {
  icon: LucideIcon;
  label: string;
}

const apqpSteps: Step[] = [
  { icon: Target, label: "Plan" },
  { icon: PenTool, label: "Design" },
  { icon: FlaskConical, label: "Test & Validate" },
  { icon: Factory, label: "Launch Production" },
];

const ppapSteps: Step[] = [
  { icon: FileSpreadsheet, label: "Parts List" },
  { icon: Route, label: "Process Steps" },
  { icon: AlertTriangle, label: "Risk Checks" },
  { icon: ClipboardCheck, label: "Control Plan" },
];

const dashedLine = {
  backgroundImage:
    "repeating-linear-gradient(to bottom, hsl(var(--border)) 0, hsl(var(--border)) 8px, transparent 8px, transparent 16px)",
};

export const WhatIsPpap = () => (
  <section className="border-b border-border bg-secondary/40 py-24">
    <div className="container">
      <div className="mx-auto mb-14 max-w-2xl text-center">
        <div className="mb-4 text-xs font-semibold uppercase tracking-[0.2em] text-primary">
          Plain English, No Jargon
        </div>
        <h2 className="text-balance text-4xl font-bold md:text-5xl">Wait — what's PPAP/APQP?</h2>
        <p className="mt-4 text-lg text-muted-foreground">
          If you're new to the auto-parts world, here's the visual 30-second version.
        </p>
      </div>

      <div className="mx-auto grid max-w-4xl gap-12 md:grid-cols-2 md:gap-8">
        {/* APQP track — the process */}
        <div>
          <Badge variant="secondary" className="mb-6">
            APQP · the process
          </Badge>
          <div className="relative pl-2">
            <div className="absolute left-[19px] top-2 bottom-2 w-px animate-flow-line-y" style={dashedLine} />
            <div className="space-y-6">
              {apqpSteps.map((step, i) => (
                <RevealNode key={step.label} index={i} className="relative flex items-center gap-4">
                  <div className="relative z-10 flex h-10 w-10 shrink-0 items-center justify-center rounded-full border-2 border-primary/40 bg-card">
                    <step.icon className="h-4 w-4 text-primary" />
                  </div>
                  <span className="text-sm font-semibold">{step.label}</span>
                </RevealNode>
              ))}
            </div>
          </div>
          <div className="mt-6 pl-2">
            <ApqpDeepDive />
          </div>
        </div>

        {/* PPAP track — the proof */}
        <div>
          <Badge className="mb-6 bg-accent text-accent-foreground hover:bg-accent-hover">PPAP · the proof</Badge>
          <div className="relative pl-2">
            <div className="absolute left-[19px] top-2 bottom-2 w-px animate-flow-line-y" style={dashedLine} />
            <div className="space-y-6">
              {ppapSteps.map((step, i) => (
                <RevealNode key={step.label} index={i} className="relative flex items-center gap-4">
                  <div className="relative z-10 flex h-10 w-10 shrink-0 items-center justify-center rounded-full border-2 border-accent/40 bg-card">
                    <step.icon className="h-4 w-4 text-accent" />
                  </div>
                  <span className="text-sm font-semibold">{step.label}</span>
                </RevealNode>
              ))}
              <RevealNode index={ppapSteps.length} className="relative flex items-center gap-4">
                <div className="relative z-10 flex h-12 w-12 shrink-0 items-center justify-center rounded-full border-2 border-accent bg-accent/10">
                  <FolderOpen className="h-5 w-5 text-accent" />
                </div>
                <div className="flex flex-wrap items-center gap-2">
                  <span className="text-sm font-semibold">Proof Package</span>
                  <span className="inline-flex -rotate-3 items-center gap-1 rounded-full border border-success/30 bg-success/10 px-2 py-0.5 text-[10px] font-bold uppercase tracking-wide text-success">
                    <CheckCircle2 className="h-3 w-3" /> Approved
                  </span>
                </div>
              </RevealNode>
            </div>
          </div>
          <div className="mt-6 pl-2">
            <PpapDeepDive />
          </div>
        </div>
      </div>

      {/* Convergence — both tracks lead to one outcome */}
      <div className="mx-auto mt-6 flex max-w-4xl flex-col items-center gap-2">
        <div className="h-8 w-px animate-flow-line-y" style={dashedLine} />
        <RevealNode index={5} className="flex items-center gap-3 rounded-full border border-border bg-card px-6 py-3 shadow-card">
          <Handshake className="h-5 w-5 text-primary" />
          <span className="text-sm font-semibold">Part accepted by the customer</span>
        </RevealNode>
      </div>
    </div>
  </section>
);
