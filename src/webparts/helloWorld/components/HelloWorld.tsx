import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react';
import { MSGraphClient } from '@microsoft/sp-http';  
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types'; 

const Dashboard: React.FunctionComponent<IHelloWorldProps> = (props) => {
  const getData = async () => 
  {
    props.context.msGraphClientFactory
    .getClient()
    .then((client: MSGraphClient) => {
      client
        .api("me/drive/special/approot:/CrossDeviceApp/settings.json:/content")
        .version("v1.0")
        .put('{"key": "value"}', (err, res) => {
          console.log(err,res);
        });
    });
  };

  return (
    <div className={styles.helloWorld}>
      <div className={styles.container}>
        <div className={styles.row}>
          <div className={styles.column}>
            <span className={styles.title}>Welcome to SharePoint!</span>
            <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
            <p className={styles.description}>{escape(props.description)}</p>
            <a href="https://aka.ms/spfx" className={styles.button}>
              <span className={styles.label}>Learn more</span>
            </a>
          </div>
        </div>
        <div className={styles.row}>
          <div className={styles.column}>
            <PrimaryButton text='Get current user data' onClick={getData} ></PrimaryButton>
          </div>
        </div>
      </div>
    </div>
  );
};

export default Dashboard;

// export default class HelloWorld extends React.Component<IHelloWorldProps, {}> {
//   public render(): React.ReactElement<IHelloWorldProps> {
//     return (
//       <div className={ styles.helloWorld }>
//         <div className={ styles.container }>
//           <div className={ styles.row }>
//             <div className={ styles.column }>
//               <span className={ styles.title }>Welcome to SharePoint!</span>
//               <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
//               <p className={ styles.description }>{escape(this.props.description)}</p>
//               <a href="https://aka.ms/spfx" className={ styles.button }>
//                 <span className={ styles.label }>Learn more</span>
//               </a>
//             </div>
//           </div>
//         </div>
//       </div>
//     );
//   }
// }
