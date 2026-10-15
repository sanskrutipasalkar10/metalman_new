import { ArrowRight, Target, PenTool, Route, FlaskConical, Repeat, type LucideIcon } from "lucide-react";
import { Button } from "@/components/ui/button";
import { DeepDiveDialog } from "@/components/landing/DeepDiveDialog";

interface Phase {
  icon: LucideIcon;
  name: string;
  desc: string;
  tieIn?: string;
}

const phases: Phase[] = [
  {
    icon: Target,
    name: "Plan & Define",
    desc: "Understand what the customer needs and whether it's feasible to build.",
  },
  {
    icon: PenTool,
    name: "Product Design & Development",
    desc: "Turn requirements into a CAD design, materials, and a first-pass risk review.",
  },
  {
    icon: Route,
    name: "Process Design & Development",
    desc: "Map out exactly how the part gets manufactured, station by station, with the right tooling.",
    tieIn: "Metalman auto-generates your Process Flow Diagram & Fixture List here",
  },
  {
    icon: FlaskConical,
    name: "Product & Process Validation",
    desc: "Run trial builds, check every dimension, confirm the process is repeatable.",
    tieIn: "Metalman auto-generates your PFMEA, Control Plan & Fitment Checksheet here",
  },
  {
    icon: Repeat,
    name: "Feedback, Assessment & Corrective Action",
    desc: "Once in full production, keep monitoring and fixing issues as they come up.",
  },
];

export const ApqpDeepDive = () => (
  <DeepDiveDialog
    trigger={
      <Button variant="outline" size="sm" className="gap-1.5">
        See all 5 phases <ArrowRight className="h-3.5 w-3.5" />
      </Button>
    }
    title="APQP — Advanced Product Quality Planning"
    subtitle="The 5 phases a manufacturer works through to take a part from idea to full production."
  >
    <div className="space-y-4 py-2">
      {phases.map((phase, i) => (
        <div key={phase.name} className="flex gap-4 rounded-xl border border-border bg-card p-5">
          <div className="flex h-11 w-11 shrink-0 items-center justify-center rounded-full border-2 border-primary/40 bg-secondary text-sm font-bold text-primary">
            {i + 1}
          </div>
          <div className="flex-1">
            <div className="mb-1 flex items-center gap-2">
              <phase.icon className="h-4 w-4 text-primary" />
              <h3 className="font-bold">{phase.name}</h3>
            </div>
            <p className="text-sm text-muted-foreground">{phase.desc}</p>
            {phase.tieIn && (
              <div className="mt-3 inline-flex items-center gap-1.5 rounded-full border border-accent/30 bg-accent/10 px-3 py-1 text-xs font-medium text-foreground">
                <span className="h-1.5 w-1.5 rounded-full bg-accent" />
                {phase.tieIn}
              </div>
            )}
          </div>
        </div>
      ))}
    </div>
  </DeepDiveDialog>
);
