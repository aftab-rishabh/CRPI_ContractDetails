import * as React from 'react';
import styles from './ContractDetailsNew.module.scss';
import { IContractDetailsNewProps } from './IContractDetailsNewProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Accordion } from '../../accordion';
export default class ContractDetailsNew extends React.Component<IContractDetailsNewProps, {}> {
  public render(): React.ReactElement<IContractDetailsNewProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.contractDetailsNew} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <Accordion title={'Project Details'}>
          </Accordion>
          <Accordion title={'Sales Team Feedback'}>
            
            </Accordion>
            <Accordion title={'Technical Team Feedback'}>
            
            </Accordion>
            <Accordion title={'Project Details'}>
            
            </Accordion>
        </div>

      </section>
    );
  }
}
