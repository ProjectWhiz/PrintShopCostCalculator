// src/components/FileUpload.js
import React, { useState } from 'react';

const FileUpload = () => {
  const [fileName, setFileName] = useState('');
  const [fileData, setFileData] = useState(null);

  const handleFileChange = (e) => {
    const file = e.target.files[0];
    if (file) {
      setFileName(file.name);

      const reader = new FileReader();
      reader.onload = (event) => {
        const text = event.target.result;
        setFileData(text);
        console.log('CSV file contents:', text); // TEMP: display content
      };
      reader.readAsText(file);
    }
  };

  return (
    <div style={{ textAlign: 'center' }}>
      <p>Select a CSV file to generate a 10% inventory sample.</p>
      <input type="file" accept=".csv" onChange={handleFileChange} />
      {fileName && <p>✅ File selected: <strong>{fileName}</strong></p>}
    </div>
  );
};

export default FileUpload;
