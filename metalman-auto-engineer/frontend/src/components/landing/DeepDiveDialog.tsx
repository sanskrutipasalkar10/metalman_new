import type { ReactNode } from "react";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
  DialogTrigger,
} from "@/components/ui/dialog";

export const DeepDiveDialog = ({
  trigger,
  title,
  subtitle,
  children,
}: {
  trigger: ReactNode;
  title: string;
  subtitle: string;
  children: ReactNode;
}) => (
  <Dialog>
    <DialogTrigger asChild>{trigger}</DialogTrigger>
    <DialogContent className="h-[90vh] w-[95vw] max-w-5xl overflow-y-auto sm:rounded-xl">
      <DialogHeader>
        <DialogTitle className="font-display text-2xl">{title}</DialogTitle>
        <DialogDescription>{subtitle}</DialogDescription>
      </DialogHeader>
      {children}
    </DialogContent>
  </Dialog>
);
