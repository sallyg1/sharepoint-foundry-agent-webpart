import * as React from 'react';
import { useState, useRef, useEffect } from 'react';
import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import type { IDallasPermitChatProps } from './IDallasPermitChatProps';

interface IMessage {
  id: string;
  sender: 'user' | 'agent';
  text: string;
  timestamp: string;
}

const SAMPLE_QUESTIONS: string[] = [
  'Do I need a permit for a fence?',
  'What are the remodel checklist requirements?',
  'Residential addition permit guidelines'
];

export default function DallasPermitChat(props: IDallasPermitChatProps): React.ReactElement {
  const [messages, setMessages] = useState<IMessage[]>([
    {
      id: 'welcome',
      sender: 'agent',
      text: `Hello ${props.userDisplayName || 'there'}! I am the City of Dallas Permit Assistant. How can I help you with your building or remodeling permits today?`,
      timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
    }
  ]);
  const [input, setInput] = useState<string>('');
  const [loading, setLoading] = useState<boolean>(false);
  const [conversationId, setConversationId] = useState<string | undefined>(undefined);
  const chatBottomRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    chatBottomRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages, loading]);

  const sendMessage = async (textToSend?: string): Promise<void> => {
    const text = (textToSend || input).trim();
    if (!text || loading) return;

    const userMessage: IMessage = {
      id: `user-${Date.now()}`,
      sender: 'user',
      text,
      timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
    };

    setMessages(prev => [...prev, userMessage]);
    setInput('');
    setLoading(true);

    try {
      // 1. Acquire Entra token for Tenant A API App (trim any accidental trailing whitespace)
      const targetApi = (props.apiAppIdUri || '').trim();
      const targetEndpoint = (props.functionEndpoint || '').trim();
      const client: AadHttpClient = await props.httpClientFactory.getClient(targetApi);

      // 2. Call Azure Function (auth broker)
      const response: HttpClientResponse = await client.post(
        targetEndpoint,
        AadHttpClient.configurations.v1,
        {
          headers: {
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({
            message: text,
            conversationId: conversationId
          })
        }
      );

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`HTTP ${response.status} (${response.statusText}): ${errorText}`);
      }

      interface IChatResponse {
        reply?: string;
        response?: string;
        conversationId?: string;
      }
      const data: IChatResponse = await response.json();
      if (data.conversationId) {
        setConversationId(data.conversationId);
      }

      const agentMessage: IMessage = {
        id: `agent-${Date.now()}`,
        sender: 'agent',
        text: data.reply || data.response || '(No reply text returned)',
        timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
      };

      setMessages(prev => [...prev, agentMessage]);
    } catch (err: unknown) {
      const errMsg = err instanceof Error ? err.message : String(err);
      setMessages(prev => [
        ...prev,
        {
          id: `error-${Date.now()}`,
          sender: 'agent',
          text: `Error contacting permit service: ${errMsg}`,
          timestamp: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
        }
      ]);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{
      maxWidth: 760,
      margin: '0 auto',
      background: '#ffffff',
      borderRadius: 12,
      border: '1px solid #e2e8f0',
      boxShadow: '0 4px 16px rgba(0, 0, 0, 0.06)',
      overflow: 'hidden',
      fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif'
    }}>
      {/* Header */}
      <div style={{
        background: '#0f172a',
        color: '#ffffff',
        padding: '16px 20px',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between'
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          <div style={{
            width: 38,
            height: 38,
            borderRadius: 8,
            background: '#0284c7',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            fontSize: 20
          }}>
            D
          </div>
          <div>
            <h3 style={{ margin: 0, fontSize: 16, fontWeight: 600 }}>City of Dallas Permit Assistant</h3>
            <p style={{ margin: 0, fontSize: 12, color: '#94a3b8' }}>
              Azure AI Foundry Agent · Cross-Cloud Bridge
            </p>
          </div>
        </div>
        <div style={{
          fontSize: 11,
          background: 'rgba(56, 189, 248, 0.15)',
          color: '#38bdf8',
          padding: '4px 8px',
          borderRadius: 4,
          border: '1px solid rgba(56, 189, 248, 0.3)'
        }}>
          Connected
        </div>
      </div>

      {/* Suggested prompts */}
      <div style={{
        background: '#f8fafc',
        padding: '10px 16px',
        borderBottom: '1px solid #e2e8f0',
        display: 'flex',
        gap: 8,
        flexWrap: 'wrap'
      }}>
        <span style={{ fontSize: 12, color: '#64748b', alignSelf: 'center', fontWeight: 500 }}>
          Try:
        </span>
        {SAMPLE_QUESTIONS.map((q, idx) => (
          <button
            key={idx}
            onClick={() => { sendMessage(q).catch(console.error); }}
            disabled={loading}
            style={{
              background: '#ffffff',
              border: '1px solid #cbd5e1',
              borderRadius: 16,
              padding: '4px 10px',
              fontSize: 11.5,
              color: '#0f172a',
              cursor: loading ? 'not-allowed' : 'pointer'
            }}
          >
            {q}
          </button>
        ))}
      </div>

      {/* Message history */}
      <div style={{
        height: 340,
        overflowY: 'auto',
        padding: 16,
        background: '#f8fafc',
        display: 'flex',
        flexDirection: 'column',
        gap: 12
      }}>
        {messages.map(m => {
          const isUser = m.sender === 'user';
          return (
            <div
              key={m.id}
              style={{
                display: 'flex',
                justifyContent: isUser ? 'flex-end' : 'flex-start'
              }}
            >
              <div style={{
                maxWidth: '78%',
                padding: '10px 14px',
                borderRadius: 10,
                background: isUser ? '#0284c7' : '#ffffff',
                color: isUser ? '#ffffff' : '#0f172a',
                border: isUser ? 'none' : '1px solid #e2e8f0',
                boxShadow: isUser ? '0 2px 6px rgba(2, 132, 199, 0.25)' : '0 1px 3px rgba(0, 0, 0, 0.04)',
                fontSize: 13.5,
                lineHeight: 1.5,
                whiteSpace: 'pre-wrap'
              }}>
                <div>{m.text}</div>
                <div style={{
                  fontSize: 10,
                  marginTop: 4,
                  textAlign: 'right',
                  opacity: 0.7
                }}>
                  {m.timestamp}
                </div>
              </div>
            </div>
          );
        })}
        {loading && (
          <div style={{ display: 'flex', justifyContent: 'flex-start' }}>
            <div style={{
              padding: '10px 14px',
              borderRadius: 10,
              background: '#ffffff',
              border: '1px solid #e2e8f0',
              color: '#64748b',
              fontSize: 13,
              fontStyle: 'italic'
            }}>
              Consulting Dallas permit guidelines...
            </div>
          </div>
        )}
        <div ref={chatBottomRef} />
      </div>

      {/* Input Footer */}
      <div style={{
        padding: 14,
        background: '#ffffff',
        borderTop: '1px solid #e2e8f0',
        display: 'flex',
        gap: 10
      }}>
        <input
          type="text"
          value={input}
          placeholder="Ask a question about permits, checklists, or zoning..."
          disabled={loading}
          onChange={e => setInput(e.target.value)}
          onKeyDown={e => {
            if (e.key === 'Enter') {
              e.preventDefault();
              sendMessage().catch(console.error);
            }
          }}
          style={{
            flex: 1,
            padding: '10px 14px',
            borderRadius: 8,
            border: '1px solid #cbd5e1',
            fontSize: 13.5,
            outline: 'none'
          }}
        />
        <button
          onClick={() => { sendMessage().catch(console.error); }}
          disabled={loading || !input.trim()}
          style={{
            background: loading || !input.trim() ? '#94a3b8' : '#0284c7',
            color: '#ffffff',
            border: 'none',
            borderRadius: 8,
            padding: '0 18px',
            fontSize: 13.5,
            fontWeight: 600,
            cursor: loading || !input.trim() ? 'not-allowed' : 'pointer'
          }}
        >
          Send
        </button>
      </div>
    </div>
  );
}
