import mammoth from "mammoth";
import XLSX from "xlsx";

const SOURCE_PATH = "quiz.docx"
const OUTPUT_PATH = "quiz_export.xlsx"


async function parseQuizDocument() {
    try {
        // Extract HTML to preserve formatting (bold text)
        const result = await mammoth.convertToHtml({ path: SOURCE_PATH });
        const htmlContent = result.value;
        
        const paragraphs = htmlContent
            .split(/<\/p>|<p[^>]*>/)
            .map(p => p.trim())
            .filter(p => p.length > 0);
        
        const lines = paragraphs.map(paragraph => {
            return paragraph
                .replace(/<strong>/g, '**BOLD_START**')
                .replace(/<\/strong>/g, '**BOLD_END**')
                .replace(/<b>/g, '**BOLD_START**')
                .replace(/<\/b>/g, '**BOLD_END**')
                .replace(/<[^>]*>/g, '') // Remove all other HTML tags
                .replace(/&nbsp;/g, ' ') // Replace non-breaking spaces
                .trim();
        }).filter(line => line.length > 0);

        const questions = [];
        
        for (let i = 0; i < lines.length; i += 5) {
            if (i + 4 < lines.length) {
                const questionData = {
                    number: Math.floor(i / 5) + 1,
                    question: lines[i].replace(/\*\*BOLD_START\*\*|\*\*BOLD_END\*\*/g, ''),
                    options: [],
                    correctAnswer: null
                };
                
                for (let j = 1; j <= 4; j++) {
                    const optionText = lines[i + j];
                    const isBold = optionText.includes('**BOLD_START**') && optionText.includes('**BOLD_END**');
                    const cleanText = optionText.replace(/\*\*BOLD_START\*\*|\*\*BOLD_END\*\*/g, '');
                    
                    questionData.options.push({
                        text: cleanText.slice(2), // remove a) b) c) d)
                        isCorrect: isBold
                    });
                    
                    if (isBold) {
                        questionData.correctAnswer = j; // Use option number (1, 2, 3, 4) instead of letter
                    }
                }
                questions.push(questionData);
            }
        }
        
        console.log(`Parsed ${questions.length} questions successfully!`);
        await exportToExcel(questions);
        
        return questions;
        
    } catch (error) {
        console.error("Error reading the docx file:", error);
    }
}

async function exportToExcel(questions) {
    try {
        const workbook = XLSX.utils.book_new();
        const mainSheetData = [];
        mainSheetData.push([
            'Question Text',
            'Question Type',
            'Option 1',
            'Option 2', 
            'Option 3',
            'Option 4',
            'Option 5',
            'Correct Answer',
            'Time in seconds'
        ]);
        
        questions.forEach(q => {
            const row = [
                q.question,                           
                'Multiple Choice',                    
                q.options[0]?.text || '',            
                q.options[1]?.text || '',            
                q.options[2]?.text || '',            
                q.options[3]?.text || '',            
                '',                                  
                q.correctAnswer || 'Not detected',   
                '30'                                 
            ];
            mainSheetData.push(row);
        });
        
        const mainSheet = XLSX.utils.aoa_to_sheet(mainSheetData);
        
        const colWidths = [
            { wch: 50 }, 
            { wch: 15 }, 
            { wch: 30 }, 
            { wch: 30 }, 
            { wch: 30 }, 
            { wch: 30 }, 
            { wch: 30 }, 
            { wch: 15 }, 
            { wch: 15 }  
        ];
        mainSheet['!cols'] = colWidths;
        
        XLSX.utils.book_append_sheet(workbook, mainSheet, "Quiz Questions");
        
        const filename = OUTPUT_PATH;
        XLSX.writeFile(workbook, filename);
        console.log(`✅ Excel file exported successfully: ${filename}`);
    } catch (error) {
        console.error("Error exporting to Excel:", error);
    }
}

parseQuizDocument();

