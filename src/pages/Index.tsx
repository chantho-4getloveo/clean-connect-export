
import { useState } from "react";
import { useToast } from "@/hooks/use-toast";
import FileUploadSection from "@/components/FileUploadSection";
import DataPreviewSection from "@/components/DataPreviewSection";
import ThemeToggle from "@/components/ThemeToggle";
import { processExcelFile } from "@/lib/excelProcessor";

interface Contact {
  CID: string | number;
  AID: string | number;
  Name: string;
  Phone: string;
  cleanedPhone?: string;
}

const Index = () => {
  const [contacts, setContacts] = useState<Contact[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [stats, setStats] = useState({
    totalProcessed: 0,
    duplicatesRemoved: 0,
    invalidPhones: 0
  });
  const { toast } = useToast();

  const handleFileUpload = async (file: File) => {
    if (!file) return;
    
    try {
      setIsProcessing(true);
      
      const { cleanedData, statistics } = await processExcelFile(file);
      
      setContacts(cleanedData);
      setStats(statistics);
      
      toast({
        title: "File Processed Successfully",
        description: `Processed ${statistics.totalProcessed} rows with ${statistics.duplicatesRemoved} duplicates removed.`,
      });
    } catch (error) {
      console.error("Error processing file:", error);
      toast({
        title: "Error Processing File",
        description: error instanceof Error ? error.message : "Failed to process the file. Please check the format and try again.",
        variant: "destructive",
      });
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="min-h-screen bg-background text-foreground">
      <header className="p-4 border-b border-border flex justify-between items-center">
        <h1 className="text-2xl font-bold">CleanConnect</h1>
        <ThemeToggle />
      </header>
      
      <main className="max-w-7xl mx-auto p-4 md:p-6 space-y-8">
        <FileUploadSection onFileUpload={handleFileUpload} isProcessing={isProcessing} />
        
        {contacts.length > 0 && (
          <DataPreviewSection 
            contacts={contacts} 
            stats={stats} 
          />
        )}
      </main>
    </div>
  );
};

export default Index;
