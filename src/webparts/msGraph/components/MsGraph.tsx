import * as React from 'react';
import styles from './MsGraph.module.scss';
import { IMsGraphProps } from './IMsGraphProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Bar } from 'react-chartjs-2';
import Chart from './Chart';
export default class MsGraph extends React.Component<IMsGraphProps, { name: any, email: any, charData: {} }> {

  constructor(props) {
    super(props);
    this.state = {
      name: '',
      email: '',
      charData:{
        labels: ['Red', 'Blue', 'Yellow', 'Green', 'Purple', 'Orange'],
        datasets: [{
          label: '# of Votes',
          data: [0, 0, 0, 0, 0, 0],
          backgroundColor: [
            'rgba(255, 99, 132, 0.2)',
            'rgba(54, 162, 235, 0.2)',
            'rgba(255, 206, 86, 0.2)',
            'rgba(75, 192, 192, 0.2)',
            'rgba(153, 102, 255, 0.2)',
            'rgba(255, 159, 64, 0.2)'
          ],
          borderColor: [
            'rgba(255, 99, 132, 1)',
            'rgba(54, 162, 235, 1)',
            'rgba(255, 206, 86, 1)',
            'rgba(75, 192, 192, 1)',
            'rgba(153, 102, 255, 1)',
            'rgba(255, 159, 64, 1)'
          ],
          borderWidth: 1
        }]
      }
    };
  }
 
  public componentDidMount(): void {
    this.props.graphClient
      .api("sites/danesh96.sharepoint.com,1e8f08be-d4db-43d8-a398-198099a9378b,a89d96e0-bb32-4cc1-a664-2d63c703214b/drive/root:/new.xlsx:/workbook/tables('1')/rows")
      .get((error: any, response: any, rawResponse?: any) => {
        if(response !== undefined){
          this.setState({
            name: '',
            email: "email",
            ...this.state,
            charData: {
              ...this.state.charData,
              labels:response.value.map(o=>o.values[0][1]),
              datasets:[{
                label: 'Oracle code',
                data:  response.value.map(o=>o.values[0][9]),
                backgroundColor: [
                  'rgba(255, 99, 132, 0.2)',
                  'rgba(54, 162, 235, 0.2)',
                  'rgba(255, 206, 86, 0.2)',
                  'rgba(75, 192, 192, 0.2)',
                  'rgba(153, 102, 255, 0.2)',
                  'rgba(255, 159, 64, 0.2)'
                ],
                borderColor: [
                  'rgba(255, 99, 132, 1)',
                  'rgba(54, 162, 235, 1)',
                  'rgba(255, 206, 86, 1)',
                  'rgba(75, 192, 192, 1)',
                  'rgba(153, 102, 255, 1)',
                  'rgba(255, 159, 64, 1)'
                ],
                borderWidth: 1
              }]
            }
          });
        }
      });
  }



  public render(): React.ReactElement<IMsGraphProps> {
    return (
      <div className={styles.msGraph}>
        Purchase order
        <Chart chartData = {this.state.charData}
        />
      </div>
    );
  }
}
