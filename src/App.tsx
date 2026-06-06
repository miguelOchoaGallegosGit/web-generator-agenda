import { CalendarDays, CheckCircle2, Download, Loader2, Sparkles } from "lucide-react";
import { useEffect, useMemo, useState } from "react";
import {
  AgendaModelId,
  agendaModels,
  diffInDays,
  formatInputDate,
  generateAgendaDocx,
  parseInputDate,
  validateDateRange,
} from "./agendaDoc";

const waitingMessages = [
  "Ordenando días, meses y pequeños pendientes...",
  "Acomodando líneas para que tu impresora sonría...",
  "Puliendo esquinas de la agenda...",
  "Convenciendo a Word de portarse elegante...",
  "Preparando páginas listas para anotar grandes planes...",
];

function App() {
  const today = useMemo(() => new Date(), []);
  const defaultEnd = useMemo(() => {
    const date = new Date(today);
    date.setDate(date.getDate() + 30);
    return date;
  }, [today]);

  const [startDate, setStartDate] = useState(formatInputDate(today));
  const [endDate, setEndDate] = useState(formatInputDate(defaultEnd));
  const [selectedModel, setSelectedModel] = useState<AgendaModelId>("clasica");
  const [isGenerating, setIsGenerating] = useState(false);
  const [messageIndex, setMessageIndex] = useState(0);
  const [successMessage, setSuccessMessage] = useState("");
  const [downloadName, setDownloadName] = useState("");
  const [downloadUrl, setDownloadUrl] = useState("");

  const validationMessage = validateDateRange(startDate, endDate);
  const rangeDays = useMemo(() => {
    if (!startDate || !endDate || validationMessage) {
      return 0;
    }

    return diffInDays(parseInputDate(startDate), parseInputDate(endDate)) + 1;
  }, [endDate, startDate, validationMessage]);

  const canGenerate = !validationMessage && !isGenerating;

  useEffect(() => {
    if (!isGenerating) {
      return undefined;
    }

    const timer = window.setInterval(() => {
      setMessageIndex((index) => (index + 1) % waitingMessages.length);
    }, 1800);

    return () => window.clearInterval(timer);
  }, [isGenerating]);

  useEffect(() => {
    return () => {
      if (downloadUrl) {
        URL.revokeObjectURL(downloadUrl);
      }
    };
  }, [downloadUrl]);

  function clearGeneratedFile() {
    setSuccessMessage("");
    setDownloadName("");
    setDownloadUrl((currentUrl) => {
      if (currentUrl) {
        URL.revokeObjectURL(currentUrl);
      }

      return "";
    });
  }

  async function handleGenerate() {
    if (!canGenerate) {
      return;
    }

    setIsGenerating(true);
    clearGeneratedFile();
    setMessageIndex(0);

    try {
      await pause(450);
      const blob = await generateAgendaDocx(startDate, endDate, selectedModel);
      const fileName = `agenda-${selectedModel}-${startDate}-al-${endDate}.docx`;
      const fileUrl = URL.createObjectURL(blob);
      setDownloadUrl(fileUrl);
      triggerDownload(fileUrl, fileName);
      setDownloadName(fileName);
      setSuccessMessage("Gracias por usar la aplicación. Tu agenda quedó lista para imprimir.");
    } catch (error) {
      setSuccessMessage("");
      window.alert("No se pudo generar el archivo. Revisa las fechas e inténtalo nuevamente.");
      console.error(error);
    } finally {
      setIsGenerating(false);
    }
  }

  return (
    <main className="app-shell">
      <section className="workspace" aria-busy={isGenerating}>
        <div className="intro-band">
          <div>
            <p className="eyebrow">
              <Sparkles size={16} aria-hidden="true" />
              Agenda imprimible
            </p>
            <h1>Generador de agendas Word</h1>
          </div>
          <div className="range-pill">
            <CalendarDays size={18} aria-hidden="true" />
            <span>{rangeDays > 0 ? `${rangeDays} días` : "Rango pendiente"}</span>
          </div>
        </div>

        <div className="tool-layout">
          <form className="control-panel" onSubmit={(event) => event.preventDefault()}>
            <div className="field-grid">
              <label>
                <span>Fecha de inicio</span>
                <input
                  type="date"
                  value={startDate}
                  disabled={isGenerating}
                  onChange={(event) => {
                    setStartDate(event.target.value);
                    clearGeneratedFile();
                  }}
                />
              </label>

              <label>
                <span>Fecha de fin</span>
                <input
                  type="date"
                  value={endDate}
                  disabled={isGenerating}
                  onChange={(event) => {
                    setEndDate(event.target.value);
                    clearGeneratedFile();
                  }}
                />
              </label>
            </div>

            {validationMessage ? (
              <p className="validation-message" role="alert">
                {validationMessage}
              </p>
            ) : (
              <p className="range-summary">
                Del {formatReadableDate(startDate)} al {formatReadableDate(endDate)}
              </p>
            )}

            <div className="model-grid" role="radiogroup" aria-label="Modelos de agenda">
              {agendaModels.map((model) => {
                const checked = selectedModel === model.id;

                return (
                  <label className={`model-card ${checked ? "is-selected" : ""}`} key={model.id}>
                    <input
                      type="radio"
                      name="agenda-model"
                      value={model.id}
                      checked={checked}
                      disabled={isGenerating}
                      onChange={() => {
                        setSelectedModel(model.id);
                        clearGeneratedFile();
                      }}
                    />
                    <img src={model.image} alt={`Vista previa ${model.label}`} />
                    <span className="model-copy">
                      <strong>{model.label}</strong>
                      <small>{model.description}</small>
                    </span>
                    {checked ? <CheckCircle2 className="check-icon" size={24} aria-hidden="true" /> : null}
                  </label>
                );
              })}
            </div>

            <button className="generate-button" type="button" disabled={!canGenerate} onClick={handleGenerate}>
              {isGenerating ? (
                <>
                  <Loader2 className="spin" size={20} aria-hidden="true" />
                  Generando
                </>
              ) : (
                <>
                  <Download size={20} aria-hidden="true" />
                  Generar Word
                </>
              )}
            </button>
          </form>

          <aside className="status-panel" aria-live="polite">
            {isGenerating ? (
              <div className="loading-state">
                <Loader2 className="spin large" size={42} aria-hidden="true" />
                <p>{waitingMessages[messageIndex]}</p>
              </div>
            ) : successMessage ? (
              <div className="success-state">
                <CheckCircle2 size={42} aria-hidden="true" />
                <p>{successMessage}</p>
                <span>{downloadName}</span>
                {downloadUrl ? (
                  <a className="download-link" href={downloadUrl} download={downloadName}>
                    <Download size={18} aria-hidden="true" />
                    Descargar nuevamente
                  </a>
                ) : null}
              </div>
            ) : (
              <div className="preview-state">
                <img
                  src={agendaModels.find((model) => model.id === selectedModel)?.image}
                  alt="Modelo seleccionado"
                />
              </div>
            )}
          </aside>
        </div>
      </section>
    </main>
  );
}

function formatReadableDate(value: string): string {
  if (!value) {
    return "";
  }

  return new Intl.DateTimeFormat("es-PE", {
    day: "2-digit",
    month: "long",
    year: "numeric",
  }).format(parseInputDate(value));
}

function pause(milliseconds: number): Promise<void> {
  return new Promise((resolve) => window.setTimeout(resolve, milliseconds));
}

function triggerDownload(url: string, fileName: string) {
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
}

export default App;
