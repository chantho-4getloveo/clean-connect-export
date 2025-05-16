
import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { useToast } from "@/hooks/use-toast";
import { Download } from "lucide-react";
import { exportToExcel } from "@/lib/excelProcessor";

interface Contact {
  CID: string | number;
  AID: string | number;
  Name: string;
  Phone: string;
  cleanedPhone?: string;
}

interface DataPreviewSectionProps {
  contacts: Contact[];
  stats: {
    totalProcessed: number;
    duplicatesRemoved: number;
    invalidPhones: number;
  };
}

const DataPreviewSection = ({ contacts, stats }: DataPreviewSectionProps) => {
  const [currentPage, setCurrentPage] = useState(1);
  const [isExporting, setIsExporting] = useState(false);
  const { toast } = useToast();
  const itemsPerPage = 10;
  
  const startIndex = (currentPage - 1) * itemsPerPage;
  const endIndex = startIndex + itemsPerPage;
  const currentContacts = contacts.slice(startIndex, endIndex);
  const totalPages = Math.ceil(contacts.length / itemsPerPage);

  const handleExport = async () => {
    try {
      setIsExporting(true);
      await exportToExcel(contacts);
      toast({
        title: "Success",
        description: "File downloaded successfully!",
      });
    } catch (error) {
      console.error("Export error:", error);
      toast({
        title: "Export Failed",
        description: "Failed to generate Excel file. Please try again.",
        variant: "destructive",
      });
    } finally {
      setIsExporting(false);
    }
  };

  return (
    <div className="space-y-4">
      <Card>
        <CardHeader>
          <CardTitle>Processing Summary</CardTitle>
        </CardHeader>
        <CardContent>
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            <div className="p-4 bg-muted rounded-lg text-center">
              <p className="text-sm text-muted-foreground">Total Rows</p>
              <p className="text-2xl font-bold">{stats.totalProcessed}</p>
            </div>
            <div className="p-4 bg-muted rounded-lg text-center">
              <p className="text-sm text-muted-foreground">Duplicates Removed</p>
              <p className="text-2xl font-bold">{stats.duplicatesRemoved}</p>
            </div>
            <div className="p-4 bg-muted rounded-lg text-center">
              <p className="text-sm text-muted-foreground">Invalid Phone Numbers</p>
              <p className="text-2xl font-bold">{stats.invalidPhones}</p>
            </div>
          </div>
          <Button 
            className="w-full mt-4 text-lg py-6"
            onClick={handleExport}
            disabled={isExporting}
          >
            <Download className="mr-2" />
            {isExporting ? "Generating File..." : "Download Cleaned File"}
          </Button>
        </CardContent>
      </Card>

      <Card>
        <CardHeader>
          <CardTitle>Data Preview</CardTitle>
        </CardHeader>
        <CardContent>
          <div className="relative overflow-x-auto">
            <table className="w-full text-sm text-left">
              <thead className="text-xs uppercase bg-muted">
                <tr>
                  <th className="px-4 py-3">CID</th>
                  <th className="px-4 py-3">AID</th>
                  <th className="px-4 py-3">Name</th>
                  <th className="px-4 py-3">Phone</th>
                </tr>
              </thead>
              <tbody>
                {currentContacts.map((contact, index) => (
                  <tr 
                    key={index} 
                    className={`border-b ${
                      contact.Phone === "" ? "bg-destructive/10" : ""
                    }`}
                  >
                    <td className="px-4 py-3">{contact.CID}</td>
                    <td className="px-4 py-3">{contact.AID}</td>
                    <td className="px-4 py-3">{contact.Name}</td>
                    <td className="px-4 py-3 font-mono">{contact.Phone}</td>
                  </tr>
                ))}
                {currentContacts.length === 0 && (
                  <tr>
                    <td colSpan={4} className="px-4 py-3 text-center">
                      No data available
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>

          {totalPages > 1 && (
            <div className="flex justify-between items-center mt-4">
              <Button
                variant="outline"
                onClick={() => setCurrentPage(prev => Math.max(prev - 1, 1))}
                disabled={currentPage === 1}
              >
                Previous
              </Button>
              <span className="text-sm">
                Page {currentPage} of {totalPages}
              </span>
              <Button
                variant="outline"
                onClick={() => setCurrentPage(prev => Math.min(prev + 1, totalPages))}
                disabled={currentPage === totalPages}
              >
                Next
              </Button>
            </div>
          )}
        </CardContent>
      </Card>
    </div>
  );
};

export default DataPreviewSection;
