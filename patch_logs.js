const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf-8');

const groupedCode = `
  const groupedLogsForReport = useMemo(() => {
    const groups: Record<string, typeof allLogs> = {};
    filteredLogsForReport.forEach(log => {
      const datePart = (log.timestamp || '').split(' - ')[0] || 'N/A';
      if (!groups[datePart]) {
        groups[datePart] = [];
      }
      groups[datePart].push(log);
    });
    
    const sortedDates = Object.keys(groups).sort((a, b) => {
      if (a === 'N/A') return 1;
      if (b === 'N/A') return -1;
      const partsA = a.split('/');
      const partsB = b.split('/');
      if (partsA.length === 3 && partsB.length === 3) {
        return new Date(\`\${partsB[2]}-\${partsB[1]}-\${partsB[0]}\`).getTime() - new Date(\`\${partsA[2]}-\${partsA[1]}-\${partsA[0]}\`).getTime();
      }
      return 0;
    });
    
    return sortedDates.map(date => ({
      date,
      logs: groups[date].sort((a, b) => {
        const timeA = (a.timestamp || '').split(' - ')[1] || '';
        const timeB = (b.timestamp || '').split(' - ')[1] || '';
        return timeB.localeCompare(timeA);
      })
    }));
  }, [filteredLogsForReport]);
`;

code = code.replace(
  "}, [allLogs, activeReportStartDate, activeReportEndDate]);",
  "}, [allLogs, activeReportStartDate, activeReportEndDate]);\n" + groupedCode
);

fs.writeFileSync('src/App.tsx', code);
