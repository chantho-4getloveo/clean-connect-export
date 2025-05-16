import * as XLSX from 'xlsx';

interface Contact {
  CID: string | number;
  AID: string | number;
  "Customer Name"?: string;
  Name?: string;
  Phone: string;
  cleanedPhone?: string;
}

interface ProcessingResult {
  cleanedData: Contact[];
  originalData: Contact[];
  statistics: {
    totalProcessed: number;
    duplicatesRemoved: number;
    invalidPhones: number;
  };
}

// Clean and standardize phone numbers
const cleanPhoneNumber = (phone: string): string => {
  if (!phone || typeof phone !== 'string') return '';

  // Split by multiple possible delimiters (comma, slash, @)
  const phoneSegments = phone.split(/[,\/]/)
    .map(p => p.trim())
    .filter(p => p.length > 0);
  
  // Set to keep track of unique cleaned numbers
  const uniqueNumbers = new Set<string>();
  
  for (const segment of phoneSegments) {
    // Skip segments that are clearly not phone numbers (like Telegram usernames)
    if (segment.toLowerCase().includes('telegram') || segment.startsWith('@')) {
      continue;
    }
    
    // Extract potential phone numbers from the segment
    const potentialNumbers = segment.split(/\s+/);
    
    for (const potentialNumber of potentialNumbers) {
      // Remove all non-numeric characters
      let cleanedNumber = potentialNumber.replace(/\D/g, '');
      
      // Skip if empty after cleaning
      if (!cleanedNumber) continue;
      
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
      const isValidLength = cleanedNumber.length >= 7 && cleanedNumber.length <= 11;
      
      if (isValidLength) {
        uniqueNumbers.add(cleanedNumber);
      }
    }
  }
  
  // Convert set back to array, add semicolons, and join with spaces
  return Array.from(uniqueNumbers).map(num => num + ';').join(' ');
};

// Create a global set to track all unique phone numbers
const globalUniquePhones = new Set<string>();

// Reset global unique phones set
export const resetGlobalUniquePhones = () => {
  globalUniquePhones.clear();
};

// Process Excel file
export const processExcelFile = async (file: File): Promise<ProcessingResult> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        // Reset global unique phones set for fresh processing
        resetGlobalUniquePhones();
        
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        
        // Assume the first sheet is the one to process
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Convert to JSON
        const jsonData = XLSX.utils.sheet_to_json<Contact>(worksheet);
        
        // Save a copy of the original data before cleaning
        const originalData = JSON.parse(JSON.stringify(jsonData));
        
        // Validate required columns
        if (jsonData.length > 0) {
          const firstRow = jsonData[0];
          const requiredColumns = ['CID', 'AID', 'Phone'];
          const nameColumn = firstRow.hasOwnProperty('Name') ? 'Name' : 'Customer Name';
          
          if (!(nameColumn in firstRow)) {
            requiredColumns.push('Name'); // For error message consistency
          }
          
          for (const column of requiredColumns) {
            if (!(column in firstRow) && !(column === 'Name' && 'Customer Name' in firstRow)) {
              reject(new Error(`Required column "${column}" is missing from the Excel file.`));
              return;
            }
          }
          
          // Normalize "Name" and "Customer Name" columns
          jsonData.forEach(contact => {
            if (contact.Name && !contact["Customer Name"]) {
              contact["Customer Name"] = contact.Name;
            }
          });
        }
        
        // First pass: clean phone numbers
        let invalidPhones = 0;
        const contactsWithCleanedPhones = jsonData.map(contact => {
          const cleanedPhone = cleanPhoneNumber(contact.Phone?.toString() || '');
          if (!cleanedPhone) invalidPhones++;
          
          return {
            ...contact,
            Phone: cleanedPhone
          };
        });
        
        // Second pass: deduplicate across all contacts
        const uniqueContacts: Contact[] = [];
        const phoneMap = new Map<string, Contact>();
        
        contactsWithCleanedPhones.forEach(contact => {
          if (!contact.Phone) {
            // Keep contacts with no phone numbers
            uniqueContacts.push(contact);
            return;
          }
          
          // Get individual numbers from the phone field (without semicolons)
          const phoneNumbers = contact.Phone.split(/\s+/)
            .map(p => p.replace(/;$/, ''))
            .filter(p => p.length > 0);
          
          // If all phone numbers have been seen before, skip this contact
          let allDuplicate = phoneNumbers.length > 0;
          for (const number of phoneNumbers) {
            if (!globalUniquePhones.has(number)) {
              allDuplicate = false;
              globalUniquePhones.add(number);
            }
          }
          
          if (!allDuplicate) {
            uniqueContacts.push(contact);
          }
        });
        
        const duplicatesRemoved = contactsWithCleanedPhones.length - uniqueContacts.length;
        
        resolve({
          cleanedData: uniqueContacts,
          originalData: originalData,
          statistics: {
            totalProcessed: originalData.length,
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
export const exportToExcel = async (cleanedContacts: Contact[], originalContacts: Contact[]): Promise<void> => {
  try {
    // Create a new workbook
    const workbook = XLSX.utils.book_new();
    
    // Prepare original data with consistent column structure
    const preparedOriginalData = originalContacts.map(contact => {
      const customerName = contact.Name || contact["Customer Name"] || "";
      return {
        CID: contact.CID,
        AID: contact.AID,
        "Customer Name": customerName,
        Phone: contact.Phone
      };
    });
    
    // Create the original data worksheet
    const originalWorksheet = XLSX.utils.json_to_sheet(preparedOriginalData);
    XLSX.utils.book_append_sheet(workbook, originalWorksheet, 'Original Data');
    
    // Prepare cleaned data with consistent column structure
    const preparedCleanedData = cleanedContacts.map(contact => {
      const customerName = contact.Name || contact["Customer Name"] || "";
      return {
        CID: contact.CID,
        AID: contact.AID,
        "Customer Name": customerName,
        Phone: contact.Phone
      };
    });
    
    // Create the cleaned data worksheet
    const cleanedWorksheet = XLSX.utils.json_to_sheet(preparedCleanedData);
    XLSX.utils.book_append_sheet(workbook, cleanedWorksheet, 'Cleaned Data');
    
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
