import * as React from 'react';
import styles from './McpDemo.module.scss';
import type { IMcpDemoProps } from './IMcpDemoProps';
import { Form } from './Form/form';

export default class McpDemo extends React.Component<IMcpDemoProps> {
  public render(): React.ReactElement<IMcpDemoProps> {
    const {
      hasTeamsContext,
      context
    } = this.props;

    return (
      <section className={`${styles.mcpDemo} ${hasTeamsContext ? styles.teams : ''} bg-gray-50/30 p-2 sm:p-6 min-h-screen`}>
        <Form context={context} />
      </section>
    );
  }
}

