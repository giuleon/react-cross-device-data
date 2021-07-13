import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  PrimaryButton,
  TextField
} from 'office-ui-fabric-react';
import { MSGraphClient } from '@microsoft/sp-http';

interface IUserData {
  Theme: string;
  Token: string;
  Preference1: string;
}

const Dashboard: React.FunctionComponent<IHelloWorldProps> = (props) => {
  const [userSettings, setUserSettings] = React.useState<IUserData | null>(null);

  const setUserData = async () => {
    props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient) => {
        client
          .api("me/drive/special/approot:/CrossDeviceApp/settings.json:/content")
          .version("v1.0")
          .put('{"Theme": "Dark", "Token": "123456789", "Preference1": "some value"}', (err, res) => {
            console.log(err, res);
          });
      });
  };

  const getUserData = async () => {
    const msGraphClient = await props.context.msGraphClientFactory.getClient();
    const result = await msGraphClient
      .api("me/drive/special/approot:/CrossDeviceApp/settings.json?select=@microsoft.graph.downloadUrl")
      .version("v1.0")
      .get();
    console.log(result);
    const response = await fetch(`${result['@microsoft.graph.downloadUrl']}`);
    console.log('response', response);
    const userData: IUserData = await response.json();
    console.log('userData', userData);
    return userData;

    // .then((client: MSGraphClient) => {
    //   client
    //     .api("me/drive/special/approot:/CrossDeviceApp/settings.json?select=@microsoft.graph.downloadUrl")
    //     .version("v1.0")
    //     .get(async (err, res) => {
    //       console.log(err,res);
    //       const response = await fetch(`${res['@microsoft.graph.downloadUrl']}`);
    //       console.log('response', response);
    //       const userData: IUserData = await response.json();
    //       console.log('userData', userData);
    //       return userData;
    //     });
    // });
  };

  const loadUserData = async () => {
    return await getUserData();
  };

  React.useEffect(() => {
    console.log('useEffect');
    (async () => {
      let data = await getUserData();
      console.log('loadUserData', data);
      setUserSettings(data);
    })();
  }, []);

  let view: any = '';
  if (userSettings !== null) {
    view =
      <div style={{backgroundColor: '#fff'}}>
        <div className={styles.row}>
          <div className={styles.column}>
            <TextField label="Standard:" underlined placeholder="Enter text here" defaultValue={userSettings.Theme} />
          </div>
        </div>
        <div className={styles.row}>
          <div className={styles.column}>
            <TextField label="Disabled:" underlined placeholder="Enter text here" defaultValue={userSettings.Token} />
          </div>
        </div>
        <div className={styles.row}>
          <div className={styles.column}>
            <TextField label="Required:" underlined placeholder="Enter text here" defaultValue={userSettings.Preference1} />
          </div>
        </div>
      </div>;
  }

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
        {view}
        <div className={styles.row}>
          <div className={styles.column}>
            <PrimaryButton text='Save current user data' onClick={setUserData} ></PrimaryButton>&#9;
            <PrimaryButton text='Get current user data' onClick={getUserData} ></PrimaryButton>
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
