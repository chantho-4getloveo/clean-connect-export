
import { useState, useCallback } from "react";
import { useDropzone } from "react-dropzone";
import { Upload, File } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";

interface FileUploadSectionProps {
  onFileUpload: (file: File) => void;
  isProcessing: boolean;
}

const FileUploadSection = ({ onFileUpload, isProcessing }: FileUploadSectionProps) => {
  const [selectedFile, setSelectedFile] = useState<File | null>(null);

  const onDrop = useCallback((acceptedFiles: File[]) => {
    const file = acceptedFiles[0];
    if (file) {
      setSelectedFile(file);
    }
  }, []);

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls'],
    },
    maxFiles: 1,
    disabled: isProcessing
  });

  const handleUpload = () => {
    if (selectedFile) {
      onFileUpload(selectedFile);
    }
  };

  return (
    <Card className="bg-card">
      <CardContent className="pt-6">
        <div
          {...getRootProps()}
          className={`border-2 border-dashed rounded-lg p-8 text-center cursor-pointer transition-colors ${
            isDragActive ? "border-primary bg-primary/5" : "border-border"
          } ${isProcessing ? "opacity-50 cursor-not-allowed" : ""}`}
        >
          <input {...getInputProps()} />
          <div className="flex flex-col items-center justify-center space-y-4">
            <Upload className="h-12 w-12 text-muted-foreground" />
            <div>
              {isDragActive ? (
                <p className="text-lg">Drop the Excel file here...</p>
              ) : (
                <div className="space-y-2">
                  <p className="text-lg font-medium">
                    Drag & drop an Excel file here, or click to select
                  </p>
                  <p className="text-sm text-muted-foreground">
                    Supported formats: .xlsx, .xls
                  </p>
                  <p className="text-sm text-muted-foreground">
                    Required columns: CID, AID, Name, Phone
                  </p>
                </div>
              )}
            </div>
          </div>
        </div>

        {selectedFile && (
          <div className="mt-4">
            <div className="flex items-center justify-between p-3 bg-muted rounded-md">
              <div className="flex items-center space-x-2">
                <File className="h-5 w-5 text-muted-foreground" />
                <span className="text-sm font-medium truncate max-w-[300px]">
                  {selectedFile.name}
                </span>
              </div>
              <Button 
                onClick={handleUpload} 
                disabled={isProcessing}
              >
                {isProcessing ? "Processing..." : "Process File"}
              </Button>
            </div>
          </div>
        )}
      </CardContent>
    </Card>
  );
};

export default FileUploadSection;
