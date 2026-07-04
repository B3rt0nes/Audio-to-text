"use client";

import React, { useState, useEffect, useRef } from 'react';
import { Mic, Send, StopCircle, User, Bot, AlertCircle } from 'lucide-react';

interface Message {
  id: string;
  text: string;
  sender: 'user' | 'bot';
  isAudio?: boolean;
}

import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

const COCA_QUESTIONS = [
  "Quali sono i paletti che direzionano la scelta tra il tempo della vita personale e quello della vita scout?",
  "Quali sono i criteri che ci guidano nel fare le staff e le cose da fare (Sostenibilità del servizio)?",
  "Devo togliere qualcosa per poterlo sostenere o esiste un altro modo per far si che abiti la mia vita in modo equilibrato: 'Conciliare o tagliare'?",
  "Racconta le tue buone pratiche."
];

export default function Home() {
  const [userName, setUserName] = useState('');
  const [isAuth, setIsAuth] = useState(false);
  const [tempName, setTempName] = useState('');
  const [currentQuestionIndex, setCurrentQuestionIndex] = useState(0);
  
  const [messages, setMessages] = useState<Message[]>([]);
  const [responsesData, setResponsesData] = useState<any[]>([]); // Memorizza le risposte
  
  const [inputText, setInputText] = useState('');
  const [isRecording, setIsRecording] = useState(false);
  const [recordingError, setRecordingError] = useState<string | null>(null);
  const [isFinished, setIsFinished] = useState(false);
  
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const recognitionRef = useRef<any>(null);

  // Send the first question when authenticated
  useEffect(() => {
    if (isAuth && messages.length === 0) {
      setMessages([
        { 
          id: '1', 
          text: `Buona caccia ${userName}! Iniziamo la riflessione sulla Sostenibilità del Servizio.\n\nEcco il primo spunto:\n**${COCA_QUESTIONS[0]}**`, 
          sender: 'bot' 
        }
      ]);
    }
  }, [isAuth, userName, messages.length]);

  // Initialize Speech Recognition
  useEffect(() => {
    if (typeof window !== 'undefined' && isAuth && !isFinished) {
      const SpeechRecognition = (window as any).SpeechRecognition || (window as any).webkitSpeechRecognition;
      if (SpeechRecognition) {
        recognitionRef.current = new SpeechRecognition();
        recognitionRef.current.continuous = true;
        recognitionRef.current.interimResults = true;
        recognitionRef.current.lang = 'it-IT'; 

        recognitionRef.current.onresult = (event: any) => {
          let interimTranscript = '';
          let finalTranscript = '';

          for (let i = event.resultIndex; i < event.results.length; ++i) {
            if (event.results[i].isFinal) {
              finalTranscript += event.results[i][0].transcript;
            } else {
              interimTranscript += event.results[i][0].transcript;
            }
          }
          
          if (finalTranscript) {
            setInputText((prev) => prev + finalTranscript + ' ');
          }
        };

        recognitionRef.current.onerror = (event: any) => {
          console.error("Speech recognition error", event.error);
          setRecordingError(`Errore nel microfono: ${event.error}`);
          setIsRecording(false);
        };
        
        recognitionRef.current.onend = () => {
          if (isRecording) {
            try {
              recognitionRef.current.start();
            } catch (e) {
              setIsRecording(false);
            }
          }
        };
      } else {
        setRecordingError('Il tuo browser non supporta la registrazione vocale.');
      }
    }
  }, [isRecording, isAuth, isFinished]);

  useEffect(() => {
    if (isAuth) {
      messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
    }
  }, [messages, isAuth]);

  const toggleRecording = () => {
    if (isRecording) {
      setIsRecording(false);
      if (recognitionRef.current) {
        recognitionRef.current.stop();
      }
      setTimeout(() => {
        if (inputText.trim()) {
          handleSend(inputText, true);
        }
      }, 500);
    } else {
      setRecordingError(null);
      setInputText(''); 
      if (recognitionRef.current) {
        try {
          recognitionRef.current.start();
          setIsRecording(true);
        } catch (e) {
          console.error(e);
          setRecordingError("Non riesco ad avviare il microfono.");
        }
      }
    }
  };

  const downloadExcel = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Riflessioni CoCa');
    
    worksheet.columns = [
      { header: 'Data e Ora', key: 'timestamp', width: 25 },
      { header: 'Mittente', key: 'sender', width: 20 },
      { header: 'Domanda', key: 'question', width: 50 },
      { header: 'Tipo Input', key: 'type', width: 20 },
      { header: 'Risposta', key: 'text', width: 100 }
    ];
    
    worksheet.getRow(1).font = { bold: true };
    
    responsesData.forEach(resp => {
      worksheet.addRow(resp);
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, `riflessioni_coca_${userName.replace(/\s+/g, '_')}.xlsx`);
  };

  const handleSend = async (text: string = inputText, isVoiceMsg = false) => {
    if (!text.trim()) return;

    const newMessage: Message = {
      id: Date.now().toString(),
      text: text.trim(),
      sender: 'user',
      isAudio: isVoiceMsg,
    };

    setMessages(prev => [...prev, newMessage]);
    setInputText('');
    
    // Salva risposta in memoria
    const currentQuestionTitle = `${currentQuestionIndex + 1}. ${COCA_QUESTIONS[currentQuestionIndex]}`;
    setResponsesData(prev => [...prev, {
      timestamp: new Date().toLocaleString('it-IT'),
      sender: userName,
      question: currentQuestionTitle,
      type: isVoiceMsg ? 'Vocale (Trascritto)' : 'Testuale',
      text: text.trim()
    }]);

    // Passa alla prossima domanda o concludi
    setTimeout(async () => {
      if (currentQuestionIndex < COCA_QUESTIONS.length - 1) {
        const nextIndex = currentQuestionIndex + 1;
        setCurrentQuestionIndex(nextIndex);
        
        const botMessage: Message = {
          id: (Date.now() + 1).toString(),
          text: `Ottimo, grazie! Passiamo al prossimo spunto:\n\n**${COCA_QUESTIONS[nextIndex]}**`,
          sender: 'bot'
        };
        setMessages(prev => [...prev, botMessage]);
      } else {
        setIsFinished(true);
        const botMessage: Message = {
          id: (Date.now() + 1).toString(),
          text: `Grazie mille per le tue riflessioni, ${userName}! Abbiamo concluso le domande sulla Sostenibilità del Servizio.\n\nBuona Strada!`,
          sender: 'bot'
        };
        setMessages(prev => [...prev, botMessage]);
      }
    }, 1000);
  };

  if (!isAuth) {
    return (
      <div className="flex flex-col items-center justify-center h-screen bg-background text-foreground font-sans px-4">
        <div className="max-w-md w-full bg-muted/30 p-8 rounded-3xl border border-border shadow-2xl backdrop-blur-sm">
          <div className="flex justify-center mb-6">
            <div className="w-16 h-16 rounded-full bg-primary/20 flex items-center justify-center text-primary">
              <User size={32} />
            </div>
          </div>
          <h1 className="text-2xl font-bold text-center mb-2">Riflessione Co.Ca.</h1>
          <p className="text-muted-foreground text-center mb-8">Inserisci il tuo nome per partecipare alla riflessione sulla Sostenibilità del Servizio.</p>
          
          <form onSubmit={(e) => {
            e.preventDefault();
            if (tempName.trim()) {
              setUserName(tempName.trim());
              setIsAuth(true);
            }
          }} className="flex flex-col gap-4">
            <input 
              type="text" 
              placeholder="Il tuo nome..." 
              value={tempName}
              onChange={(e) => setTempName(e.target.value)}
              className="w-full py-3 px-4 rounded-xl border border-border bg-background focus:outline-none focus:ring-2 focus:ring-primary/50 transition-all text-center text-lg"
              autoFocus
            />
            <button 
              type="submit"
              disabled={!tempName.trim()}
              className="w-full py-3 rounded-xl bg-primary text-primary-foreground font-semibold shadow-lg shadow-primary/30 hover:bg-primary/90 disabled:opacity-50 disabled:cursor-not-allowed transition-all"
            >
              Inizia la Riflessione
            </button>
          </form>
        </div>
      </div>
    );
  }

  // Progress percentage
  const progressPercent = Math.min(100, Math.round(((currentQuestionIndex + (isFinished ? 1 : 0)) / COCA_QUESTIONS.length) * 100));

  return (
    <div className="flex flex-col h-screen bg-background text-foreground font-sans">
      
      {/* Header with Progress Bar */}
      <header className="px-6 py-4 border-b border-border bg-background/50 backdrop-blur-md sticky top-0 z-10 flex flex-col gap-3">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-full bg-primary/20 flex items-center justify-center text-primary">
              <Bot size={24} />
            </div>
            <div>
              <h1 className="text-xl font-semibold leading-tight">Sostenibilità del Servizio</h1>
              <p className="text-sm text-muted-foreground">Loggato come <span className="font-semibold text-primary">{userName}</span> (Co.Ca.)</p>
            </div>
          </div>
          <button 
            onClick={() => { setIsAuth(false); setUserName(''); setCurrentQuestionIndex(0); setMessages([]); setIsFinished(false); }}
            className="text-sm text-muted-foreground hover:text-foreground transition-colors"
          >
            Esci
          </button>
        </div>
        {/* Progress Bar */}
        <div className="w-full h-2 bg-muted rounded-full overflow-hidden">
           <div 
             className="h-full bg-primary transition-all duration-500 ease-out" 
             style={{ width: `${progressPercent}%` }}
           ></div>
        </div>
        <div className="text-xs text-muted-foreground text-right font-medium">
           {isFinished ? 'Completato' : `Spunto ${currentQuestionIndex + 1} di ${COCA_QUESTIONS.length}`}
        </div>
      </header>

      {/* Chat Area */}
      <main className="flex-1 overflow-y-auto p-6 space-y-6 flex flex-col">
        {messages.map((msg) => (
          <div key={msg.id} className={`flex gap-3 max-w-[85%] ${msg.sender === 'user' ? 'self-end flex-row-reverse' : 'self-start'}`}>
            <div className={`w-10 h-10 rounded-full flex items-center justify-center flex-shrink-0 mt-1 shadow-sm ${msg.sender === 'user' ? 'bg-primary text-primary-foreground' : 'bg-secondary text-secondary-foreground border border-border'}`}>
              {msg.sender === 'user' ? <User size={18} /> : <Bot size={18} />}
            </div>
            <div className={`p-4 rounded-3xl ${
              msg.sender === 'user' 
                ? 'bg-primary text-primary-foreground rounded-tr-sm shadow-md' 
                : 'bg-muted border border-border text-foreground rounded-tl-sm shadow-md text-[1.05rem]'
            }`}>
              {msg.isAudio && (
                <div className="flex items-center gap-2 mb-2 text-primary-foreground/80 text-xs font-semibold uppercase tracking-wider">
                  <Mic size={14} /> Trascrizione Audio
                </div>
              )}
              {/* Simple markdown bold parsing for bot messages */}
              <p className="leading-relaxed whitespace-pre-wrap">
                {msg.text.split('**').map((part, i) => i % 2 === 1 ? <strong key={i} className="font-bold text-primary">{part}</strong> : part)}
              </p>
            </div>
          </div>
        ))}
        <div ref={messagesEndRef} />
      </main>

      {/* Input Area */}
      {!isFinished ? (
        <footer className="p-4 bg-background border-t border-border">
          {recordingError && (
            <div className="mb-3 px-4 py-3 bg-red-500/10 border border-red-500/20 text-red-500 rounded-xl flex items-center gap-2 text-sm">
              <AlertCircle size={16} /> {recordingError}
            </div>
          )}
          <div className="max-w-4xl mx-auto relative flex items-center gap-2">
            
            <div className="relative flex-1">
              <input 
                type="text"
                value={inputText}
                onChange={(e) => setInputText(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' && !isRecording) handleSend(inputText, false);
                }}
                placeholder={isRecording ? "In ascolto..." : "Scrivi o registra un audio per questa riflessione..."}
                disabled={isRecording}
                className={`w-full py-4 pl-6 pr-12 rounded-full border border-border bg-muted/50 focus:outline-none focus:ring-2 focus:ring-primary/50 transition-all ${
                  isRecording ? 'border-red-500/50 bg-red-500/5 text-red-500 placeholder:text-red-500/70 shadow-inner' : ''
                }`}
              />
              {isRecording && (
                 <div className="absolute right-4 top-1/2 -translate-y-1/2 flex items-center gap-2">
                    <span className="text-xs text-red-500 font-medium animate-pulse">Registrazione in corso...</span>
                    <div className="w-2 h-2 rounded-full bg-red-500 animate-pulse-ring relative"></div>
                 </div>
              )}
            </div>

            <button 
              onClick={toggleRecording}
              className={`p-4 rounded-full flex items-center justify-center transition-all ${
                isRecording 
                  ? 'bg-red-500 text-white shadow-lg shadow-red-500/30 hover:bg-red-600 scale-105' 
                  : 'bg-muted text-foreground hover:bg-border'
              }`}
              title={isRecording ? "Ferma registrazione" : "Registra audio"}
            >
              {isRecording ? <StopCircle size={22} /> : <Mic size={22} />}
            </button>

            {!isRecording && (
              <button 
                onClick={() => handleSend(inputText, false)}
                disabled={!inputText.trim()}
                className={`p-4 rounded-full flex items-center justify-center transition-all ${
                  inputText.trim() 
                    ? 'bg-primary text-primary-foreground shadow-lg shadow-primary/30 hover:bg-primary/90' 
                    : 'bg-muted text-muted-foreground cursor-not-allowed'
                }`}
                title="Invia riflessione"
              >
                <Send size={22} />
              </button>
            )}

          </div>
        </footer>
      ) : (
        <footer className="p-6 bg-background border-t border-border flex flex-col items-center gap-4">
          <p className="text-muted-foreground text-sm font-medium">
            Hai completato tutte le riflessioni previste.
          </p>
          <button 
            onClick={downloadExcel}
            className="px-6 py-3 bg-green-600 text-white rounded-xl font-bold shadow-lg shadow-green-600/30 hover:bg-green-700 transition-all flex items-center gap-2"
          >
            Scarica File Excel (Risposte)
          </button>
        </footer>
      )}
      
    </div>
  );
}
