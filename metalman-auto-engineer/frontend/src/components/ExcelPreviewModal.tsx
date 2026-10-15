import { useEffect, useState } from "react";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Loader2, Download, Maximize2, FileSpreadsheet } from "lucide-react";
import { Button } from "@/components/ui/button";
import { API_BASE } from "@/lib/api";
import { toast } from "sonner";

interface Props {
  filename: string | null;
  onClose: () => void;
}

export const ExcelPreviewModal = ({ filename, onClose }: Props) => {
  const actualFilename = filename?.split('/').pop() || "";
  const previewUrl = `${API_BASE}/api/preview/${actualFilename}`;
  const [opening, setOpening] = useState(false);

  const handleDownload = () => {
    if (filename) {
        window.open(`${API_BASE}/outputs/${actualFilename}`, '_blank');
    }
  };

  const handleEditInExcel = async () => {
    if (!actualFilename) return;
    setOpening(true);
    try {
      const resp = await fetch(`${API_BASE}/api/edit/${actualFilename}`, { method: "POST" });
      if (!resp.ok) {
        const data = await resp.json().catch(() => ({}));
        throw new Error(data.detail || "Could not open Excel.");
      }
      toast.success(`Opening ${actualFilename} in Excel on the server machine…`);
    } catch (err: any) {
      toast.error(err.message || "Could not open Excel.");
    } finally {
      setOpening(false);
    }
  };

  return (
    <Dialog open={!!filename} onOpenChange={(open) => !open && onClose()}>
      <DialogContent className="max-w-[98vw] w-[1400px] max-h-[96vh] h-[96vh] overflow-hidden flex flex-col p-0 border-none shadow-2xl">
        <DialogHeader className="p-4 border-b bg-white">
          <div className="flex items-center justify-between pr-8">
            <DialogTitle className="flex items-center gap-2 text-xl font-bold">
              <Maximize2 className="h-6 w-6 text-primary" />
              Document Preview: <span className="font-mono text-sm font-normal text-muted-foreground">{actualFilename}</span>
            </DialogTitle>
            <div className="flex items-center gap-3">
              <p className="text-[10px] uppercase tracking-widest text-muted-foreground font-bold bg-secondary px-2 py-1 rounded">High Fidelity Preview</p>
              <Button size="sm" variant="outline" onClick={handleEditInExcel} disabled={opening} className="h-9 border-success/30 text-success hover:bg-success/5 font-bold">
                {opening ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <FileSpreadsheet className="mr-2 h-4 w-4" />} Edit in Excel
              </Button>
              <Button size="sm" variant="outline" onClick={handleDownload} className="h-9 border-primary/30 text-primary hover:bg-primary/5 font-bold">
                <Download className="mr-2 h-4 w-4" /> Download Excel (.xlsx)
              </Button>
            </div>
          </div>
        </DialogHeader>

        <div className="flex-1 bg-slate-200 flex flex-col">
          <iframe 
            src={previewUrl}
            className="w-full h-full border-none bg-white"
            title="Excel Preview"
          />
        </div>
      </DialogContent>
    </Dialog>
  );
};
