import { useState, useEffect } from "react";
import { ArrowLeft, Save, LayoutDashboard, History, Settings, Bell, Download, Edit, Trash2, Crop } from "lucide-react";
import { useNavigate, useSearchParams } from "react-router-dom";
import { toast } from "sonner";
import { Button } from "@/components/ui/button";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue, SelectGroup, SelectLabel } from "@/components/ui/select";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Logo } from "@/components/Logo";
import { Avatar, AvatarFallback } from "@/components/ui/avatar";
import { Eye } from "lucide-react";
import { ExcelPreviewModal } from "@/components/ExcelPreviewModal";
import { API_BASE } from "@/lib/api";

const DOCUMENT_OPTIONS = [
  { value: "bom", label: "Bill of Materials (BOM)" },
  { value: "pfd", label: "Process Flow Diagram (PFD)" },
  { value: "tooling", label: "Tooling List" },
  { value: "fitment", label: "Fitment Checksheet" },
  { value: "pfmea_assembly", label: "PFMEA - Assembly" },
  { value: "pfmea_sheetmetal", label: "PFMEA - Sheet Metal" },
  { value: "pfmea_bop", label: "PFMEA - BOP Parts" },
  { value: "control_plan", label: "Control Plan (All Sheets)" },
  { value: "fixture_plan", label: "Fixture PM Master Plan" },
];

type ColumnOption = { value: string; label: string };
type ColumnGroup = { groupLabel: string; options: ColumnOption[] };
type ColumnConfig = (ColumnOption | ColumnGroup)[];

const COLUMN_OPTIONS: Record<string, ColumnConfig> = {
  bom: [
    { value: "S.NO", label: "S.NO" },
    { value: "Level", label: "Level" },
    { value: "Part No", label: "Part No" },
    { value: "SAP CODE", label: "SAP CODE" },
    { value: "Part Name", label: "Part Name" },

    { value: "Photo", label: "Photo" },
    { value: "Rev", label: "Rev" },
    { value: "Qyt", label: "Qyt" },
    { value: "Remarks", label: "Remarks" },
    { value: "THK", label: "THK" },
    { value: "GRADE", label: "GRADE" },
    { value: "Manufacturing Process Details", label: "Manufacturing Process Details" },
  ],
  pfd: [
    {
      groupLabel: "Sub Assmbly Index",
      options: [
        { value: "pfd_1_Index", label: "Index" },
        { value: "pfd_1_Operation", label: "Operation" },
        { value: "pfd_1_Part_Name", label: "Part Name" },
        { value: "pfd_1_Part_No", label: "Part No." },
        { value: "pfd_1_Rev", label: "Rev." },
        { value: "pfd_1_Assembly_Pic", label: "Assembly Pic" },
        { value: "pfd_1_Remarks", label: "Remarks" },
      ]
    },
    {
      groupLabel: "SUB_ASSY",
      options: [
        { value: "pfd_2_Operation_No", label: "Operation No" },
        { value: "pfd_2_OPERATION", label: "OPERATION" },
        { value: "pfd_2_MOVEMENT", label: "MOVEMENT" },
        { value: "pfd_2_INSPECTION", label: "INSPECTION" },
        { value: "pfd_2_STORE", label: "STORE" },
        { value: "pfd_2_Operation_Description", label: "Operation Description" },
        { value: "pfd_2_Machine", label: "Machine, Device, Jig, Tools for mfg" },
        { value: "pfd_2_List_of_Parts", label: "List of Parts and Sub Assemblies" },
        { value: "pfd_2_Evaluation", label: "Evaluation and Analysis Methods" },
      ]
    },
    {
      groupLabel: "Q-SHEETMETAL_PARTS",
      options: [
        { value: "pfd_3_Operation_No", label: "Operation No" },
        { value: "pfd_3_OPERATION", label: "OPERATION" },
        { value: "pfd_3_MOVEMENT", label: "MOVEMENT" },
        { value: "pfd_3_INSPECTION", label: "INSPECTION" },
        { value: "pfd_3_STORE", label: "STORE" },
        { value: "pfd_3_Operation_Description", label: "Operation Description" },
        { value: "pfd_3_Machine", label: "Machine, Device, Jig, Tools for mfg" },
        { value: "pfd_3_List_of_Parts", label: "List of Parts and Sub Assemblies" },
        { value: "pfd_3_Evaluation", label: "Evaluation and Analysis Methods" },
      ]
    },
    {
      groupLabel: "T - BOP & Hardwares",
      options: [
        { value: "pfd_4_Operation_No", label: "Operation No" },
        { value: "pfd_4_OPERATION", label: "OPERATION" },
        { value: "pfd_4_MOVEMENT", label: "MOVEMENT" },
        { value: "pfd_4_INSPECTION", label: "INSPECTION" },
        { value: "pfd_4_STORE", label: "STORE" },
        { value: "pfd_4_Operation_Description", label: "Operation Description" },
        { value: "pfd_4_Machine", label: "Machine, Device, Jig, Tools for mfg" },
        { value: "pfd_4_T_BOP", label: "T - BOP & Hardwares" },
      ]
    }
  ],
  tooling: [
    { value: "Sr.No.", label: "Sr.No." },
    { value: "Part Name", label: "Part Name" },
    { value: "Fixture No.", label: "Fixture No." },
    { value: "Operation No.", label: "Operation No." },
    { value: "Cell No.", label: "Cell No." },
    { value: "Fixture Name", label: "Fixture Name" },
    { value: "Photo", label: "Photo" },
    { value: "Welding Source", label: "Welding Source" },
    { value: "Make", label: "Make" },
    { value: "Qty", label: "Qty" },
  ],
  fitment: [
    {
      groupLabel: "Header Fields",
      options: [
        { value: "C3", label: "Line / Category (C3)" },
        { value: "C4", label: "Model (C4)" },
        { value: "C5", label: "Part Description (C5)" },
        { value: "C6", label: "Part Number — Header (C6)" },
        { value: "J4", label: "Date (J4)" },
      ]
    },
    {
      groupLabel: "Data Row Fields",
      options: [
        { value: "A", label: "Sr. No. (Col A)" },
        { value: "B", label: "Part Name (Col B)" },
        { value: "C", label: "Part No. (Col C)" },
        { value: "E", label: "Qty (Col E)" },
      ]
    },
  ],
  pfmea_assembly: [
    {
      groupLabel: "PROTOTYPE (A-D)",
      options: [
        { value: "A", label: "No. (A)" },
        { value: "B", label: "Process Function (B)" },
        { value: "C", label: "Potential Failure Mode (C)" },
        { value: "D", label: "Potential Effect(s) of Failure (D)" },
      ]
    },
    {
      groupLabel: "PRE-LAUNCH (E-L)",
      options: [
        { value: "E", label: "Severity (E)" },
        { value: "F", label: "CLA SS (F)" },
        { value: "G", label: "Potential Cause(s) of Failure (G)" },
        { value: "H", label: "Occurrence (H)" },
        { value: "I", label: "Current Controls Prevention (I)" },
        { value: "J", label: "Current Controls Detection (J)" },
        { value: "K", label: "Detection (K)" },
        { value: "L", label: "RPN (L)" },
      ]
    },
    {
      groupLabel: "PRODUCTION (M-S)",
      options: [
        { value: "M", label: "Recommended Action (M)" },
        { value: "N", label: "Responsibility & Completion Date (N)" },
        { value: "O", label: "Action Taken (O)" },
        { value: "P", label: "Revised Severity (P)" },
        { value: "Q", label: "Revised Occurrence (Q)" },
        { value: "R", label: "Revised Detection (R)" },
        { value: "S", label: "Revised RPN (S)" },
      ]
    }
  ],
  pfmea_sheetmetal: [
    {
      groupLabel: "PROTOTYPE (A-D)",
      options: [
        { value: "A", label: "No. (A)" },
        { value: "B", label: "Process Function (B)" },
        { value: "C", label: "Potential Failure Mode (C)" },
        { value: "D", label: "Potential Effect(s) of Failure (D)" },
      ]
    },
    {
      groupLabel: "PRE-LAUNCH (E-L)",
      options: [
        { value: "E", label: "Severity (E)" },
        { value: "F", label: "CLA SS (F)" },
        { value: "G", label: "Potential Cause(s) of Failure (G)" },
        { value: "H", label: "Occurrence (H)" },
        { value: "I", label: "Current Controls Prevention (I)" },
        { value: "J", label: "Current Controls Detection (J)" },
        { value: "K", label: "Detection (K)" },
        { value: "L", label: "RPN (L)" },
      ]
    },
    {
      groupLabel: "PRODUCTION (M-S)",
      options: [
        { value: "M", label: "Recommended Action (M)" },
        { value: "N", label: "Responsibility & Completion Date (N)" },
        { value: "O", label: "Action Taken (O)" },
        { value: "P", label: "Revised Severity (P)" },
        { value: "Q", label: "Revised Occurrence (Q)" },
        { value: "R", label: "Revised Detection (R)" },
        { value: "S", label: "Revised RPN (S)" },
      ]
    }
  ],
  pfmea_bop: [
    {
      groupLabel: "PROTOTYPE (A-D)",
      options: [
        { value: "A", label: "No. (A)" },
        { value: "B", label: "Process Function (B)" },
        { value: "C", label: "Potential Failure Mode (C)" },
        { value: "D", label: "Potential Effect(s) of Failure (D)" },
      ]
    },
    {
      groupLabel: "PRE-LAUNCH (E-L)",
      options: [
        { value: "E", label: "Severity (E)" },
        { value: "F", label: "CLA SS (F)" },
        { value: "G", label: "Potential Cause(s) of Failure (G)" },
        { value: "H", label: "Occurrence (H)" },
        { value: "I", label: "Current Controls Prevention (I)" },
        { value: "J", label: "Current Controls Detection (J)" },
        { value: "K", label: "Detection (K)" },
        { value: "L", label: "RPN (L)" },
      ]
    },
    {
      groupLabel: "PRODUCTION (M-S)",
      options: [
        { value: "M", label: "Recommended Action (M)" },
        { value: "N", label: "Responsibility & Completion Date (N)" },
        { value: "O", label: "Action Taken (O)" },
        { value: "P", label: "Revised Severity (P)" },
        { value: "Q", label: "Revised Occurrence (Q)" },
        { value: "R", label: "Revised Detection (R)" },
        { value: "S", label: "Revised RPN (S)" },
      ]
    }
  ],
  control_plan: [
    {
      groupLabel: "Sub Assembly Index",
      options: [
        { value: "cp_1_A", label: "Index (A)" },
        { value: "cp_1_B", label: "Operation (B)" },
        { value: "cp_1_C", label: "Part Name (C)" },
        { value: "cp_1_D", label: "Part No. (D)" },
        { value: "cp_1_E", label: "Rev. (E)" },
        { value: "cp_1_F", label: "Assembly Pic (F)" },
        { value: "cp_1_G", label: "Remarks (G)" },
      ]
    },
    {
      groupLabel: "ASSY_SUB_ASSY",
      options: [
        { value: "cp_2_A", label: "Part process/number (A)" },
        { value: "cp_2_B", label: "Process Name / Op Desc (B)" },
        { value: "cp_2_C", label: "Machine, Device, Jig, Tools (C)" },
        { value: "cp_2_D", label: "Char - No (D)" },
        { value: "cp_2_E", label: "Char - Product (E)" },
        { value: "cp_2_F", label: "Char - Process (F)" },
        { value: "cp_2_G", label: "Special Characteristic (G)" },
        { value: "cp_2_H", label: "Product/Process Spec/Tolerance (H)" },
        { value: "cp_2_I", label: "Evaluation/Measurement (I)" },
        { value: "cp_2_J", label: "Sample - Size (J)" },
        { value: "cp_2_K", label: "Sample - Freq (K)" },
        { value: "cp_2_L", label: "Control Method (L)" },
        { value: "cp_2_M", label: "Reaction Plan (M)" },
        { value: "cp_2_N", label: "Responsibility (N)" },
      ]
    },
    {
      groupLabel: "Q-Sheetmetal Parts",
      options: [
        { value: "cp_3_A", label: "Part process/number (A)" },
        { value: "cp_3_B", label: "Process Name / Op Desc (B)" },
        { value: "cp_3_C", label: "Machine, Device, Jig, Tools (C)" },
        { value: "cp_3_D", label: "Char - No (D)" },
        { value: "cp_3_E", label: "Char - Product (E)" },
        { value: "cp_3_F", label: "Char - Process (F)" },
        { value: "cp_3_G", label: "Special Characteristic (G)" },
        { value: "cp_3_H", label: "Product/Process Spec/Tolerance (H)" },
        { value: "cp_3_I", label: "Evaluation/Measurement (I)" },
        { value: "cp_3_J", label: "Sample - Size (J)" },
        { value: "cp_3_K", label: "Sample - Freq (K)" },
        { value: "cp_3_L", label: "Control Method (L)" },
        { value: "cp_3_M", label: "Reaction Plan (M)" },
        { value: "cp_3_N", label: "Responsibility (N)" },
      ]
    },
    {
      groupLabel: "T - BOP & Hardwares",
      options: [
        { value: "cp_4_A", label: "Part process/number (A)" },
        { value: "cp_4_B", label: "Process Name / Op Desc (B)" },
        { value: "cp_4_C", label: "Machine, Device, Jig, Tools (C)" },
        { value: "cp_4_D", label: "Char - No (D)" },
        { value: "cp_4_E", label: "Char - Product (E)" },
        { value: "cp_4_F", label: "Char - Process (F)" },
        { value: "cp_4_G", label: "Special Characteristic (G)" },
        { value: "cp_4_H", label: "Product/Process Spec/Tolerance (H)" },
        { value: "cp_4_I", label: "Evaluation/Measurement (I)" },
        { value: "cp_4_J", label: "Sample - Size (J)" },
        { value: "cp_4_K", label: "Sample - Freq (K)" },
        { value: "cp_4_L", label: "Control Method (L)" },
        { value: "cp_4_M", label: "Reaction Plan (M)" },
        { value: "cp_4_N", label: "Responsibility (N)" },
      ]
    }
  ],
  fixture_plan: [
    { value: "Sr.No.", label: "Sr.No. (Col A)" },
    { value: "Part Name", label: "Part Name (Col B)" },
    { value: "Fixture No.", label: "Fixture No. (Col C)" },
    { value: "OP No", label: "OP No (Col D)" },
    { value: "Fixture Name", label: "Fixture Name (Col E)" },
  ],
};
type CorrectionItem = {
  id: string;
  document: string;
  column: string;
  cellNo: string;
  replacementContent: string;
};

const Correction = () => {
  const navigate = useNavigate();
  const [searchParams] = useSearchParams();
  const taskId = searchParams.get("taskId");

  const [selectedDoc, setSelectedDoc] = useState<string>("");
  const [selectedColumn, setSelectedColumn] = useState<string>("");
  const [cellNo, setCellNo] = useState<string>("");
  const [replacementContent, setReplacementContent] = useState<string>("");

  const [corrections, setCorrections] = useState<CorrectionItem[]>([]);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [modifiedFiles, setModifiedFiles] = useState<{ document: string, url: string }[]>([]);
  const [previewFile, setPreviewFile] = useState<string | null>(null);
  const [taskFiles, setTaskFiles] = useState<{ name: string, url: string, type: string }[]>([]);

  useEffect(() => {
    if (taskId) {
      fetch(`${API_BASE}/api/status/${taskId}`)
        .then(res => res.json())
        .then(data => {
          if (data.files) {
            setTaskFiles(data.files);
          }
        })
        .catch(err => console.error("Failed to fetch task status:", err));
    }
  }, [taskId]);

  const getFileForDoc = (docType: string) => {
    // 1. Check if we have a freshly modified version first
    const modified = modifiedFiles.find(f => f.document === docType);
    if (modified) return modified.url;

    // 2. Otherwise use the original task files
    const mapping: Record<string, string> = {
      bom: "Bill of Materials",
      pfd: "Engineering PFD",
      tooling: "Tooling Master List",
      fitment: "Fitment Checksheet",
      pfmea_assembly: "Process FMEA",
      pfmea_sheetmetal: "Process FMEA",
      pfmea_bop: "Process FMEA",
      control_plan: "Control Plan",
    };

    const targetName = mapping[docType];
    if (!targetName) return null;

    return taskFiles.find(f => f.name.includes(targetName))?.url || null;
  };

  const handleDocChange = (val: string) => {
    setSelectedDoc(val);
    setSelectedColumn("");
  };

  const handleAddChange = (e: React.FormEvent) => {
    e.preventDefault();
    if (!selectedDoc || !selectedColumn || !cellNo || !replacementContent) return;

    setCorrections(prev => [
      ...prev,
      {
        id: crypto.randomUUID(),
        document: selectedDoc,
        column: selectedColumn,
        cellNo,
        replacementContent
      }
    ]);

    setCellNo("");
    setReplacementContent("");
    // We intentionally leave selectedDoc and selectedColumn so the user can keep editing the same document easily.
  };

  const handleRemoveChange = (id: string) => {
    setCorrections(prev => prev.filter(c => c.id !== id));
  };

  const handleBulkSubmit = async () => {
    if (!taskId) {
      toast.error("No Task ID found in the URL. Please launch correction from the Dashboard.");
      return;
    }
    if (corrections.length === 0) {
      toast.error("No changes to submit.");
      return;
    }

    setIsSubmitting(true);
    setModifiedFiles([]);

    try {
      const resp = await fetch(`${API_BASE}/api/correct`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          task_id: taskId,
          corrections: corrections.map(c => ({
            document: c.document,
            column: c.column,
            cell_no: c.cellNo,
            replacement_content: c.replacementContent
          }))
        }),
      });

      const data = await resp.json();
      if (!resp.ok) {
        throw new Error(data.detail || "Failed to apply corrections.");
      }

      toast.success(data.message || "All corrections applied successfully!");
      setCorrections([]);
      if (data.modified_files) {
        setModifiedFiles(data.modified_files);
      }
    } catch (err: any) {
      console.error(err);
      toast.error(err.message || "An error occurred while applying corrections.");
    } finally {
      setIsSubmitting(false);
    }
  };

  const getDocumentLabel = (docValue: string) => {
    return DOCUMENT_OPTIONS.find(d => d.value === docValue)?.label || docValue;
  };

  return (
    <div className="min-h-screen bg-secondary/30 pb-20">
      <nav className="sticky top-0 z-50 border-b border-border bg-card/80 backdrop-blur-md">
        <div className="container flex h-16 items-center justify-between">
          <div className="flex items-center gap-8">
            <Logo className="h-8" />
            <div className="hidden h-8 w-px bg-border md:block" />
            <nav className="hidden items-center gap-6 md:flex">
              <a href="/dashboard" className="flex items-center gap-2 text-sm font-medium text-muted-foreground transition-colors hover:text-foreground">
                <LayoutDashboard className="h-4 w-4" /> Dashboard
              </a>
              <a href="/cropper" className="flex items-center gap-2 text-sm font-medium text-muted-foreground transition-colors hover:text-foreground">
                <Crop className="h-4 w-4" /> Cropper
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
            <Avatar className="h-9 w-9 border border-border shadow-sm">
              <AvatarFallback className="bg-primary/5 text-xs font-bold">ME</AvatarFallback>
            </Avatar>
          </div>
        </div>
      </nav>

      <main className="container py-10">
        <div className="mx-auto max-w-5xl">
          <div className="mb-6 flex items-center justify-between">
            <div>
              <h1 className="text-3xl font-bold tracking-tight">Document Corrections</h1>
              <p className="text-muted-foreground mt-2">
                Queue up multiple precise changes across your generated documents.
              </p>
            </div>
            <Button variant="outline" onClick={() => navigate("/dashboard")} className="gap-2">
              <ArrowLeft className="h-4 w-4" /> Back to Dashboard
            </Button>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
            {/* Input Form Column */}
            <div className="lg:col-span-5 space-y-6">
              <div className="overflow-hidden rounded-2xl border border-border bg-card shadow-card animate-in fade-in slide-in-from-bottom-4 duration-500">
                <div className="p-7">
                  <h2 className="text-xl font-bold mb-6">New Change</h2>
                  <form onSubmit={handleAddChange} className="space-y-6">

                    <div className="space-y-3">
                      <div className="flex items-center justify-between">
                        <Label htmlFor="document" className="text-sm font-semibold">1. Select Document</Label>
                        {selectedDoc && getFileForDoc(selectedDoc) && (
                          <Button
                            type="button"
                            variant="link"
                            size="sm"
                            className="h-auto p-0 text-primary font-bold flex items-center gap-1 hover:no-underline"
                            onClick={() => setPreviewFile(getFileForDoc(selectedDoc))}
                          >
                            <Eye className="h-4 w-4" /> View Current Document
                          </Button>
                        )}
                      </div>
                      <Select value={selectedDoc} onValueChange={handleDocChange}>
                        <SelectTrigger className="h-11 w-full">
                          <SelectValue placeholder="Select a document type..." />
                        </SelectTrigger>
                        <SelectContent>
                          {DOCUMENT_OPTIONS.map(doc => (
                            <SelectItem key={doc.value} value={doc.value}>
                              {doc.label}
                            </SelectItem>
                          ))}
                        </SelectContent>
                      </Select>
                    </div>

                    {selectedDoc && (
                      <div className="space-y-3 animate-in fade-in duration-300">
                        <Label htmlFor="column" className="text-sm font-semibold">2. Select Target Column/Sheet</Label>
                        <Select value={selectedColumn} onValueChange={setSelectedColumn}>
                          <SelectTrigger className="h-11 w-full">
                            <SelectValue placeholder="Select the column..." />
                          </SelectTrigger>
                          <SelectContent>
                            {COLUMN_OPTIONS[selectedDoc]?.map((col, idx) => {
                              if ('groupLabel' in col) {
                                return (
                                  <SelectGroup key={idx}>
                                    <SelectLabel className="font-bold text-primary">{col.groupLabel}</SelectLabel>
                                    {col.options.map(opt => (
                                      <SelectItem key={opt.value} value={opt.value}>
                                        {opt.label}
                                      </SelectItem>
                                    ))}
                                  </SelectGroup>
                                );
                              }
                              return (
                                <SelectItem key={col.value} value={col.value}>
                                  {col.label}
                                </SelectItem>
                              );
                            })}
                          </SelectContent>
                        </Select>
                      </div>
                    )}

                    <div className="space-y-3">
                      <Label htmlFor="cellNo" className="text-sm font-semibold">3. Cell Number</Label>
                      <Input
                        id="cellNo"
                        placeholder="e.g., C4, D12"
                        className="h-11"
                        value={cellNo}
                        onChange={(e) => setCellNo(e.target.value.toUpperCase())}
                        required
                      />
                    </div>

                    <div className="space-y-3">
                      <Label htmlFor="replacementContent" className="text-sm font-semibold">4. Replacement Content</Label>
                      <Input
                        id="replacementContent"
                        placeholder="Enter new value..."
                        className="h-11"
                        value={replacementContent}
                        onChange={(e) => setReplacementContent(e.target.value)}
                        required
                      />
                    </div>

                    <div className="pt-4 border-t border-border mt-8">
                      <Button
                        type="submit"
                        variant="secondary"
                        className="h-12 w-full font-bold border border-border hover:bg-secondary/80 transition-all duration-300"
                        disabled={!selectedDoc || !selectedColumn || !cellNo || !replacementContent}
                      >
                        + Add to Queue
                      </Button>
                    </div>
                  </form>
                </div>
              </div>
            </div>

            {/* Staging Area / Queue Column */}
            <div className="lg:col-span-7 space-y-6">
              <div className="overflow-hidden rounded-2xl border border-border bg-card shadow-card flex flex-col h-full animate-in fade-in slide-in-from-bottom-4 duration-500 delay-100">
                <div className="p-7 border-b border-border bg-secondary/10 flex items-center justify-between">
                  <h2 className="text-xl font-bold flex items-center gap-2">
                    Pending Changes
                    <span className="flex items-center justify-center bg-primary text-primary-foreground text-xs rounded-full h-6 w-6">
                      {corrections.length}
                    </span>
                  </h2>
                </div>

                <div className="flex-1 p-7 overflow-auto min-h-[300px]">
                  {corrections.length === 0 ? (
                    <div className="h-full flex flex-col items-center justify-center text-muted-foreground py-10">
                      <div className="h-16 w-16 rounded-full bg-secondary flex items-center justify-center mb-4">
                        <Edit className="h-6 w-6 opacity-50" />
                      </div>
                      <p>Your queue is empty.</p>
                      <p className="text-sm">Use the form to stage your edits.</p>
                    </div>
                  ) : (
                    <div className="space-y-3">
                      {corrections.map((c, index) => (
                        <div key={c.id} className="p-4 rounded-xl border border-border bg-background flex items-center justify-between group hover:border-primary/50 transition-colors">
                          <div className="flex flex-col gap-1">
                            <div className="flex items-center gap-2 text-sm font-medium">
                              <span className="text-primary">{getDocumentLabel(c.document)}</span>
                              <span className="text-muted-foreground">•</span>
                              <span>Cell {c.cellNo}</span>
                            </div>
                            <div className="text-sm text-muted-foreground">
                              Change to: <span className="text-foreground font-semibold">"{c.replacementContent}"</span>
                            </div>
                          </div>
                          <div className="flex items-center gap-1">
                            <Button
                              variant="ghost"
                              size="icon"
                              className="text-muted-foreground hover:text-primary"
                              onClick={() => setPreviewFile(getFileForDoc(c.document))}
                              title="View Document"
                            >
                              <Eye className="h-4 w-4" />
                            </Button>
                            <Button
                              variant="ghost"
                              size="icon"
                              className="text-muted-foreground hover:text-destructive hover:bg-destructive/10"
                              onClick={() => handleRemoveChange(c.id)}
                            >
                              <Trash2 className="h-4 w-4" />
                            </Button>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>

                <div className="p-7 border-t border-border bg-secondary/10">
                  <Button
                    className="h-14 w-full text-base font-bold bg-accent text-accent-foreground shadow-orange hover:translate-y-[-2px] hover:shadow-orange-lg transition-all duration-300"
                    disabled={corrections.length === 0 || isSubmitting}
                    onClick={handleBulkSubmit}
                  >
                    {isSubmitting ? (
                      <span className="flex items-center gap-2">
                        <Save className="mr-2 h-5 w-5 animate-pulse" /> Processing...
                      </span>
                    ) : (
                      <>
                        <Save className="mr-2 h-5 w-5" /> Submit All Changes
                      </>
                    )}
                  </Button>
                </div>
              </div>

              {/* Download Buttons Area */}
              {modifiedFiles.length > 0 && (
                <div className="space-y-3 animate-in fade-in slide-in-from-bottom-2 duration-300">
                  <h3 className="text-sm font-bold text-muted-foreground uppercase tracking-wider mb-3">Updated Files Available</h3>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                    {modifiedFiles.map((file, idx) => (
                      <div key={idx} className="flex gap-2">
                        <Button
                          type="button"
                          variant="outline"
                          className="h-14 flex-1 text-sm font-bold border-2 border-primary text-primary hover:bg-primary/5 transition-all duration-300 justify-start px-4"
                          onClick={() => window.open(`${API_BASE}${file.url}`, "_blank")}
                        >
                          <Download className="mr-3 h-5 w-5 shrink-0" />
                          <span className="truncate text-left">
                            Download {getDocumentLabel(file.document)}
                          </span>
                        </Button>
                        <Button
                          variant="secondary"
                          className="h-14 w-14 border-2 border-border p-0"
                          onClick={() => setPreviewFile(file.url)}
                          title="Preview"
                        >
                          <Eye className="h-5 w-5" />
                        </Button>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>

          </div>
        </div>
      </main>

      <ExcelPreviewModal
        filename={previewFile}
        onClose={() => setPreviewFile(null)}
      />
    </div>
  );
};

export default Correction;
