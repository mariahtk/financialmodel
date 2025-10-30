import { useState } from "react";

export default function Home() {
  const [file, setFile] = useState(null);
  const [downloadUrl, setDownloadUrl] = useState("");

  const handleUpload = async () => {
    const formData = new FormData();
    formData.append("file", file);

    const res = await fetch("/api/run_model", {  // matches API file name
      method: "POST",
      body: formData,
    });

    if (!res.ok) return alert("Failed to process the model");

    const blob = await res.blob();
    const url = window.URL.createObjectURL(blob);
    setDownloadUrl(url);
  };

  return (
    <div style={{ padding: "2rem" }}>
      <h1>Financial Model Web App</h1>
      <input type="file" onChange={(e) => setFile(e.target.files[0])} />
      <button onClick={handleUpload} disabled={!file}>Run Model</button>
      {downloadUrl && (
        <div>
          <a href={downloadUrl} download="Processed_Model.xlsm">
            Download Processed Model
          </a>
        </div>
      )}
    </div>
  );
}
