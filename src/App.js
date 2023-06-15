import "./styles.css";
import { Renderer } from "xlsx-renderer";
import { saveAs } from "file-saver";
import viewModel from "./Data";

export default function App() {
  const handleClick = () => {
    fetch("./template.xlsx")
      .then((r) => r.blob())
      .then((xlsxBlob) => {
        const reader = new FileReader();
        reader.readAsArrayBuffer(xlsxBlob);

        reader.addEventListener("loadend", async (e) => {
          const renderer = new Renderer();

          const templateFileBuffer = reader.result;

          const result = await renderer.renderFromArrayBuffer(
            templateFileBuffer,
            viewModel
          );

          await result.xlsx
            .writeBuffer()
            .then((buffer) =>
              saveAs(new Blob([buffer]), `${Date.now()}_result_report.xlsx`)
            )
            .catch((err) => console.log("Error writing excel export", err));
        });
      });
  };

  return (
    <div className="App">
      <h1>XLSX Renderer</h1>
      <h2>Get XLSX Generated based on template provided</h2>
      <button onClick={handleClick}>Export</button>
    </div>
  );
}
