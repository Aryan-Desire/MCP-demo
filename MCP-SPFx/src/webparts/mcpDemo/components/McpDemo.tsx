import * as React from 'react';
import styles from './McpDemo.module.scss';
import type { IMcpDemoProps } from './IMcpDemoProps';
import { Form } from './Form/form';
import { ChatBot } from './ChatBot/chatBot';

export interface IMcpDemoState {
  activeTab: 'form' | 'chatbot';
}

export default class McpDemo extends React.Component<IMcpDemoProps, IMcpDemoState> {
  constructor(props: IMcpDemoProps) {
    super(props);
    this.state = {
      activeTab: 'form'
    };
  }

  public render(): React.ReactElement<IMcpDemoProps> {
    const {
      hasTeamsContext,
      context
    } = this.props;
    const { activeTab } = this.state;

    return (
      <section className={`${styles.mcpDemo} ${hasTeamsContext ? styles.teams : ''} bg-gray-50/30 p-2 sm:p-6 min-h-screen font-sans`}>
        {/* Navigation Tabs */}
        <div className="max-w-4xl mx-auto mb-6 bg-white border border-gray-200 rounded-lg p-1.5 flex space-x-1 shadow-xs">
          <button
            onClick={() => this.setState({ activeTab: 'form' })}
            className={`flex-1 py-2.5 px-4 text-sm font-semibold rounded-md transition-all duration-200 ${
              activeTab === 'form'
                ? 'bg-indigo-600 text-white shadow-xs'
                : 'text-gray-600 hover:text-gray-900 hover:bg-gray-50'
            }`}
          >
            <div className="flex items-center justify-center space-x-2">
              <svg className="w-4 h-4" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
              </svg>
              <span>School Entry Form</span>
            </div>
          </button>
          
          <button
            onClick={() => this.setState({ activeTab: 'chatbot' })}
            className={`flex-1 py-2.5 px-4 text-sm font-semibold rounded-md transition-all duration-200 ${
              activeTab === 'chatbot'
                ? 'bg-indigo-600 text-white shadow-xs'
                : 'text-gray-600 hover:text-gray-900 hover:bg-gray-50'
            }`}
          >
            <div className="flex items-center justify-center space-x-2">
              <svg className="w-4 h-4" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" d="M8 10h.01M12 10h.01M16 10h.01M21 12c0 4.418-4.03 8-9 8a9.863 9.863 0 01-4.255-.949L3 20l1.395-3.72C3.512 15.042 3 13.574 3 12c0-4.418 4.03-8 9-8s9 3.582 9 8z" />
              </svg>
              <span>SharePoint AI Chatbot</span>
            </div>
          </button>
        </div>

        {activeTab === 'form' ? (
          <Form context={context} />
        ) : (
          <ChatBot context={context} />
        )}
      </section>
    );
  }
}

