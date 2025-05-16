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

  // Split by multiple possible delimiters (comma, slash)
  const phoneNumbers = phone.split(/[,\/]/)
    .map(p => p.trim())
    .filter(p => p.length > 0);
  
  // Set to keep track of unique cleaned numbers
  const uniqueNumbers = new Set<string>();
  
  const cleanedNumbers = phoneNumbers.map(number => {
    // Remove all non-numeric characters
    let cleanedNumber = number.replace(/\D/g, '');
    
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
    
    return isValidLength ? cleanedNumber : '';
  }).filter(n => n.length > 0); // Remove any empty strings
  
  // Remove duplicates within the same contact
  cleanedNumbers.forEach(num => uniqueNumbers.add(num));
  
  // Convert set back to array, add semicolons, and join with spaces
  return Array.from(uniqueNumbers).map(num => num + ';').join(' ');
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
        
        // Save a copy of the original data before cleaning
        const originalData = JSON.parse(JSON.stringify(jsonData));
        
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
          originalData: originalData,
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
export const exportToExcel = async (cleanedContacts: Contact[], originalContacts: Contact[]): Promise<void> => {
  try {
    // Create a new workbook
    const workbook = XLSX.utils.book_new();
    
    // Rename "Name" column to "Customer Name" in original data
    const renamedOriginalContacts = originalContacts.map(contact => {
      return {
        CID: contact.CID,
        AID: contact.AID,
        "Customer Name": contact.Name,
        Phone: contact.Phone
      };
    });
    
    // Create the original data worksheet
    const originalWorksheet = XLSX.utils.json_to_sheet(renamedOriginalContacts);
    XLSX.utils.book_append_sheet(workbook, originalWorksheet, 'Original Data');
    
    // Prepare cleaned data with renamed "Name" column
    const processedContacts = cleanedContacts.map(contact => {
      return {
        CID: contact.CID,
        AID: contact.AID,
        "Customer Name": contact.Name,
        Phone: contact.Phone
      };
    });
    
    // Create the cleaned data worksheet
    const cleanedWorksheet = XLSX.utils.json_to_sheet(processedContacts);
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
