import * as React from 'react';
import styles from './MsGraph.module.scss';
import { IMsGraphProps, IMsGraphState } from './IMsGraphProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class MsGraph extends React.Component<IMsGraphProps, { name: any; email: any }> {

  constructor(props: IMsGraphProps, state: IMsGraphState) {
    super(props);
    this.state = {
      name: '',
      email: ''
    };
  }



  public componentDidMount(): void {
    this.props.graphClient
      .api('me')
      .get((error: any, user: any, rawResponse?: any) => {
        this.setState({
          name: user.displayName,
          email: user.mail
        });
      });
  }



  public render(): React.ReactElement<IMsGraphProps> {
    return (
      <div className={styles.msGraph}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!!!</span>
              <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
              <p className={styles.description}>{escape(this.props.description)}:{this.state.name}</p>
              <a href="https://aka.ms/spfx" className={styles.button}>
                <span className={styles.label}>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
