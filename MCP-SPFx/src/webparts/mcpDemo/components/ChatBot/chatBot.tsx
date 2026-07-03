import * as React from 'react';
import { useState, useRef, useEffect } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IChatBotProps {
  context: WebPartContext;
}

export interface IMessage {
  id: string;
  sender: 'user' | 'bot';
  text: string;
  timestamp: Date;
}

export const ChatBot: React.FC<IChatBotProps> = ({ context }) => {
  const [messages, setMessages] = useState<IMessage[]>([
    {
      id: 'welcome',
      sender: 'bot',
      text: 'Hello! I am your SharePoint MCP AI Assistant. Ask me anything about the lists, files, or tasks on your SharePoint sites.',
      timestamp: new Date()
    }
  ]);
  const [inputQuery, setInputQuery] = useState<string>('');
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

  const messagesEndRef = useRef<HTMLDivElement>(null);

  // Auto-scroll to bottom on new messages
  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages, isLoading]);

  const presetQueries = [
    {
      label: 'Tasks Count (ProjectManagementPortal)',
      query: 'there are how many tasks in Tasks list in https://desireinfowebsp.sharepoint.com/sites/ProjectManagementPortal site'
    },
    {
      label: 'List Schools Items',
      query: 'List all items from the Schools list on https://desireinfowebsp.sharepoint.com/sites/AIDevelopment'
    },
    {
      label: 'Available Lists',
      query: 'What lists exist in the https://desireinfowebsp.sharepoint.com/sites/ProjectManagementPortal site?'
    }
  ];

  const handleSend = async (queryToSend?: string): Promise<void> => {
    const query = (queryToSend || inputQuery).trim();
    if (!query) return;

    if (!queryToSend) {
      setInputQuery('');
    }
    setErrorMessage(null);

    // Add user message
    const userMsg: IMessage = {
      id: Math.random().toString(),
      sender: 'user',
      text: query,
      timestamp: new Date()
    };
    setMessages(prev => [...prev, userMsg]);
    setIsLoading(true);

    try {
      const response = await fetch('http://localhost:7071/api/mcpDemo', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ query })
      });

      if (response.ok) {
        const data = await response.json();
        const botMsg: IMessage = {
          id: Math.random().toString(),
          sender: 'bot',
          text: data.summary || data.message || 'No details returned from server.',
          timestamp: new Date()
        };
        setMessages(prev => [...prev, botMsg]);
      } else {
        throw new Error(`Server returned code ${response.status}: ${response.statusText}`);
      }
    } catch (err: any) {
      console.error('Error calling chatbot Azure Function:', err);
      setErrorMessage(err.message || 'Failed to get response from AI server. Please make sure the Azure Function is running locally.');
    } finally {
      setIsLoading(false);
    }
  };

  const handleKeyPress = (e: React.KeyboardEvent<HTMLInputElement>): void => {
    if (e.key === 'Enter') {
      handleSend();
    }
  };

  return (
    <div className="max-w-4xl mx-auto my-6 bg-white border border-gray-200 rounded-lg shadow-sm font-sans text-gray-800 flex flex-col h-[600px]">
      {/* Header */}
      <div className="flex items-center justify-between px-6 py-4 border-b border-gray-200 bg-gray-50/50 rounded-t-lg">
        <div className="flex items-center space-x-3">
          <div className="w-9 h-9 rounded-full bg-indigo-100 flex items-center justify-center text-indigo-600">
            <svg className="w-5 h-5" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" d="M8 10h.01M12 10h.01M16 10h.01M21 12c0 4.418-4.03 8-9 8a9.863 9.863 0 01-4.255-.949L3 20l1.395-3.72C3.512 15.042 3 13.574 3 12c0-4.418 4.03-8 9-8s9 3.582 9 8z" />
            </svg>
          </div>
          <div>
            <h2 className="text-base font-semibold text-gray-900">SharePoint AI Assistant</h2>
            <p className="text-xs text-gray-500">Connected to local SharePoint MCP</p>
          </div>
        </div>
        {isLoading && (
          <span className="text-xs text-indigo-600 font-medium animate-pulse flex items-center bg-indigo-50 px-2.5 py-1 rounded-full border border-indigo-100">
            <svg className="animate-spin h-3.5 w-3.5 mr-1.5 text-indigo-600" fill="none" viewBox="0 0 24 24">
              <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
              <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" />
            </svg>
            AI is thinking...
          </span>
        )}
      </div>

      {/* Messages Window */}
      <div className="flex-1 overflow-y-auto p-6 space-y-4 bg-slate-50/30">
        {messages.map(msg => (
          <div
            key={msg.id}
            className={`flex ${msg.sender === 'user' ? 'justify-end' : 'justify-start'}`}
          >
            <div
              className={`max-w-[80%] rounded-lg px-4 py-3 shadow-xs ${
                msg.sender === 'user'
                  ? 'bg-indigo-600 text-white rounded-br-none'
                  : 'bg-white border border-gray-200 text-gray-800 rounded-bl-none'
              }`}
            >
              {msg.sender === 'user' ? (
                <p className="text-sm whitespace-pre-wrap">{msg.text}</p>
              ) : (
                <FormattedMessage text={msg.text} />
              )}
              <span
                className={`text-[10px] block mt-1.5 ${
                  msg.sender === 'user' ? 'text-indigo-200 text-right' : 'text-gray-400'
                }`}
              >
                {msg.timestamp.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
              </span>
            </div>
          </div>
        ))}

        {/* Loading Indicator */}
        {isLoading && (
          <div className="flex justify-start">
            <div className="bg-white border border-gray-200 rounded-lg rounded-bl-none px-4 py-3 shadow-xs">
              <div className="flex items-center space-x-1 py-1">
                <div className="w-2.5 h-2.5 bg-gray-400 rounded-full animate-bounce" style={{ animationDelay: '0ms' }} />
                <div className="w-2.5 h-2.5 bg-gray-400 rounded-full animate-bounce" style={{ animationDelay: '150ms' }} />
                <div className="w-2.5 h-2.5 bg-gray-400 rounded-full animate-bounce" style={{ animationDelay: '300ms' }} />
              </div>
            </div>
          </div>
        )}

        {/* Error Banner */}
        {errorMessage && (
          <div className="p-4 bg-rose-50 border border-rose-200 rounded-lg text-rose-800 flex items-start space-x-3">
            <svg className="w-5 h-5 text-rose-500 mt-0.5 flex-shrink-0" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
            </svg>
            <div className="text-sm">
              <p className="font-semibold">Query Failed</p>
              <p className="text-xs mt-1 text-rose-700">{errorMessage}</p>
            </div>
          </div>
        )}

        <div ref={messagesEndRef} />
      </div>

      {/* Preset Suggestions */}
      <div className="px-6 py-3 border-t border-gray-100 bg-white">
        <p className="text-[11px] font-semibold text-gray-400 uppercase tracking-wider mb-2">Preset Queries</p>
        <div className="flex flex-wrap gap-2">
          {presetQueries.map((preset, idx) => (
            <button
              key={idx}
              onClick={() => handleSend(preset.query)}
              disabled={isLoading}
              className="text-xs bg-indigo-50 hover:bg-indigo-100 text-indigo-700 font-medium px-3 py-1.5 rounded-full border border-indigo-100/80 transition-all duration-150 disabled:opacity-60 disabled:pointer-events-none hover:shadow-xs"
            >
              {preset.label}
            </button>
          ))}
        </div>
      </div>

      {/* Input Form */}
      <div className="p-4 border-t border-gray-200 bg-gray-50/50 flex space-x-3 rounded-b-lg">
        <input
          type="text"
          placeholder="Ask a query about SharePoint (e.g. How many items in schools list?)..."
          value={inputQuery}
          onChange={(e) => setInputQuery(e.target.value)}
          onKeyPress={handleKeyPress}
          disabled={isLoading}
          className="flex-1 px-4 py-2.5 border border-gray-300 rounded-lg focus:outline-none focus:border-indigo-600 focus:ring-1 focus:ring-indigo-600 text-sm bg-white disabled:opacity-65"
        />
        <button
          onClick={() => handleSend()}
          disabled={isLoading || !inputQuery.trim()}
          className="px-5 py-2.5 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 font-medium transition-colors text-sm disabled:opacity-50 flex items-center justify-center shadow-sm"
        >
          <svg className="w-4 h-4 mr-1.5" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" d="M14 5l7 7m0 0l-7 7m7-7H3" />
          </svg>
          Send
        </button>
      </div>
    </div>
  );
};

// --- Custom Markdown Parser Component ---
const FormattedMessage: React.FC<{ text: string }> = ({ text }) => {
  const blocks = text.split('\n');
  return (
    <div className="space-y-2 text-sm leading-relaxed text-gray-800">
      {blocks.map((block, index) => {
        const trimmed = block.trim();
        if (!trimmed) return null;

        // Headers
        const headerMatch = trimmed.match(/^(#{1,6})\s+(.*)$/);
        if (headerMatch) {
          const level = headerMatch[1].length;
          const content = formatTextInline(headerMatch[2]);
          if (level === 1) return <h1 key={index} className="text-base font-bold text-gray-900 mt-3 mb-1">{content}</h1>;
          if (level === 2) return <h2 key={index} className="text-sm font-bold text-gray-900 mt-2.5 mb-1">{content}</h2>;
          return <h3 key={index} className="text-xs font-bold text-gray-900 mt-2 mb-0.5">{content}</h3>;
        }

        // Bullet point
        const bulletMatch = trimmed.match(/^[-*]\s+(.*)$/);
        if (bulletMatch) {
          return (
            <div key={index} className="flex items-start space-x-2 pl-2">
              <span className="text-indigo-500 select-none mt-1">•</span>
              <span className="flex-1">{formatTextInline(bulletMatch[1])}</span>
            </div>
          );
        }

        // Numbered list item
        const numberMatch = trimmed.match(/^(\d+)\.\s+(.*)$/);
        if (numberMatch) {
          return (
            <div key={index} className="flex items-start space-x-2 pl-2">
              <span className="text-indigo-500 font-medium min-w-[0.85rem] mt-0.5">{numberMatch[1]}.</span>
              <span className="flex-1">{formatTextInline(numberMatch[2])}</span>
            </div>
          );
        }

        // Code block
        if (trimmed.startsWith('```')) {
          const codeContent = trimmed.replace(/^```[a-zA-Z]*\n?/, '').replace(/```$/, '');
          return (
            <pre key={index} className="bg-gray-900 text-slate-100 p-2.5 rounded font-mono text-xs overflow-x-auto my-1.5 max-w-full">
              {codeContent}
            </pre>
          );
        }

        // Paragraph
        return <p key={index}>{formatTextInline(trimmed)}</p>;
      })}
    </div>
  );
};

// Helper for bold, code, links inline formatting
function formatTextInline(text: string): React.ReactNode[] {
  const parts: React.ReactNode[] = [];
  let remaining = text;
  let keyIdx = 0;

  while (remaining.length > 0) {
    const boldIdx = remaining.indexOf('**');
    const codeIdx = remaining.indexOf('`');
    const linkMatch = remaining.match(/\[([^\]]+)\]\(([^)]+)\)/);
    const linkIdx = linkMatch ? remaining.indexOf(linkMatch[0]) : -1;

    let closestType: 'bold' | 'code' | 'link' | 'none' = 'none';
    let minIdx = remaining.length;

    if (boldIdx !== -1 && boldIdx < minIdx) {
      minIdx = boldIdx;
      closestType = 'bold';
    }
    if (codeIdx !== -1 && codeIdx < minIdx) {
      minIdx = codeIdx;
      closestType = 'code';
    }
    if (linkIdx !== -1 && linkIdx < minIdx) {
      minIdx = linkIdx;
      closestType = 'link';
    }

    if (closestType === 'none') {
      parts.push(<span key={keyIdx++}>{remaining}</span>);
      break;
    }

    if (minIdx > 0) {
      parts.push(<span key={keyIdx++}>{remaining.substring(0, minIdx)}</span>);
    }

    remaining = remaining.substring(minIdx);

    if (closestType === 'bold') {
      const nextBold = remaining.indexOf('**', 2);
      if (nextBold !== -1) {
        const boldText = remaining.substring(2, nextBold);
        parts.push(<strong key={keyIdx++} className="font-bold text-gray-900">{boldText}</strong>);
        remaining = remaining.substring(nextBold + 2);
      } else {
        parts.push(<span key={keyIdx++}>**</span>);
        remaining = remaining.substring(2);
      }
    } else if (closestType === 'code') {
      const nextCode = remaining.indexOf('`', 1);
      if (nextCode !== -1) {
        const codeText = remaining.substring(1, nextCode);
        parts.push(<code key={keyIdx++} className="bg-gray-100 text-rose-600 px-1 py-0.5 rounded font-mono text-xs">{codeText}</code>);
        remaining = remaining.substring(nextCode + 1);
      } else {
        parts.push(<span key={keyIdx++}>`</span>);
        remaining = remaining.substring(1);
      }
    } else if (closestType === 'link' && linkMatch) {
      const display = linkMatch[1];
      const url = linkMatch[2];
      parts.push(
        <a key={keyIdx++} href={url} target="_blank" rel="noopener noreferrer" className="text-indigo-600 hover:text-indigo-800 underline transition-colors">
          {display}
        </a>
      );
      remaining = remaining.substring(linkMatch[0].length);
    }
  }

  return parts;
}
