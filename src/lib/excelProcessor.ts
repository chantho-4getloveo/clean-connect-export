
import * as XLSX from 'xlsx';

interface Contact {
  CID: string | number;
  AID: string | number;
  Name: string;
  Phone: string;
  cleanedPhone?: string;
}

interface ProcessingResult {
  cleanedData: Contact[];
  statistics: {
    totalProcessed: number;
    duplicatesRemoved: number;
    invalidPhones: number;
  };
}

// Clean and standardize phone numbers
const cleanPhoneNumber = (phone: string): string => {
  if (!phone || typeof phone !== 'string') return '';

  // Handle multiple phone numbers by taking the first one
  const firstPhone = phone.split(',')[0];
  
  // Remove all non-numeric characters
  let cleanedNumber = firstPhone.replace(/\D/g, '');
  
  // Remove country codes (assuming they start with + or 00)
  if (cleanedNumber.startsWith('00')) {
    cleanedNumber = cleanedNumber.substring(2);
  }
  if (cleanedNumber.length > 8 && cleanedNumber.startsWith('855')) {
    cleanedNumber = cleanedNumber.substring(3);
  }
  
  // If the first digit is not 0, add it (for local format)
  if (cleanedNumber.length > 0 && !cleanedNumber.startsWith('0')) {
    cleanedNumber = '0' + cleanedNumber;
  }
  
  // Check if the number has a valid length after cleaning
  // For this example, we'll consider 7-11 digits as valid
  const isValidLength = cleanedNumber.length >= 7 && cleanedNumber.length <= 11;
  
  if (!isValidLength) {
    return '';
  }
  
  // Format as :0xxxxxxxx;
  return `:${cleanedNumber};`;
};

// Remove duplicates based on cleaned phone numbers
const removeDuplicates = (data: Contact[]): Contact[] => {
  const uniquePhones = new Set<string>();
  const result: Contact[] = [];
  
  for (const contact of data) {
    // Skip duplicate check for empty phones
    if (!contact.Phone) {
      result.push(contact);
      continue;
    }
    
    if (!uniquePhones.has(contact.Phone)) {
      uniquePhones.add(contact.Phone);
      result.push(contact);
    }
  }
  
  return result;
};

// Process Excel file
export const processExcelFile = async (file: File): Promise<ProcessingResult> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        
        // Assume the first sheet is the one to process
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Convert to JSON
        const jsonData = XLSX.utils.sheet_to_json<Contact>(worksheet);
        
        // Validate required columns
        if (jsonData.length > 0) {
          const firstRow = jsonData[0];
          const requiredColumns = ['CID', 'AID', 'Name', 'Phone'];
          
          for (const column of requiredColumns) {
            if (!(column in firstRow)) {
              reject(new Error(`Required column "${column}" is missing from the Excel file.`));
              return;
            }
          }
        }
        
        // Clean phone numbers
        let invalidPhones = 0;
        const contactsWithCleanedPhones = jsonData.map(contact => {
          const cleanedPhone = cleanPhoneNumber(contact.Phone?.toString() || '');
          if (!cleanedPhone) invalidPhones++;
          
          return {
            ...contact,
            Phone: cleanedPhone
          };
        });
        
        // Remove duplicates
        const originalCount = contactsWithCleanedPhones.length;
        const uniqueContacts = removeDuplicates(contactsWithCleanedPhones);
        const duplicatesRemoved = originalCount - uniqueContacts.length;
        
        resolve({
          cleanedData: uniqueContacts,
          statistics: {
            totalProcessed: originalCount,
            duplicatesRemoved,
            invalidPhones
          }
        });
        
      } catch (error) {
        console.error('Error processing Excel file:', error);
        reject(new Error('Failed to process the Excel file. Please ensure it is a valid Excel file.'));
      }
    };
    
    reader.onerror = () => {
      reject(new Error('Error reading the file. Please try again.'));
    };
    
    reader.readAsArrayBuffer(file);
  });
};

// Export cleaned data to Excel file
export const exportToExcel = async (contacts: Contact[]): Promise<void> => {
  try {
    // Create worksheet
    const worksheet = XLSX.utils.json_to_sheet(contacts);
    
    // Create workbook
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Cleaned Contacts');
    
    // Generate file name with timestamp
    const now = new Date();
    const timestamp = now.toISOString().replace(/[-:]/g, '').split('.')[0].replace('T', '_');
    const fileName = `cleaned_contacts_${timestamp}.xlsx`;
    
    // Generate Excel file
    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const data = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    
    // Create download link
    const url = window.URL.createObjectURL(data);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    
    // Trigger download
    document.body.appendChild(link);
    link.click();
    
    // Cleanup
    window.URL.revokeObjectURL(url);
    document.body.removeChild(link);
  } catch (error) {
    console.error('Error exporting to Excel:', error);
    throw new Error('Failed to generate Excel file. Please try again.');
  }
};
