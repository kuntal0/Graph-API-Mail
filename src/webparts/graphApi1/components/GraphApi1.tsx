import * as React from 'react';
//import styles from './GraphApi1.module.scss';
import { IGraphApi1Props } from './IGraphApi1Props';
import { MSGraphClientV3 } from "@microsoft/sp-http";
import styles from './GraphApi1.module.scss';

// Email Interface
interface IEmails {
  subject: string;
  webLink: string;
  from: {
    emailAddress: {
      name: string;
      address: string;
    };
  };
  receivedDateTime: any;
  bodyPreview: string;
  isRead: any;
}

// All Items Interface
interface IAllItems {
  AllEmails: IEmails[];
}
export default class GraphApi1 extends React.Component<IGraphApi1Props, IAllItems> {
  constructor(props: IGraphApi1Props, state: IAllItems) {
    super(props);
    this.state = {
      AllEmails: []
    };
  }

  componentDidMount(): void {
    this.getMyEmails();
  }
  public getMyEmails() {
    //console.log("test emails");
    //alert("hi") 
    this.props.context.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3): void => {
        client
          .api("/me/messages")
          .version("v1.0")
          //.select("subject,weblink,from,recievedDateTime,isRead,bodyPreview")
          .get((err: any, res: any) => {
            this.setState({
              AllEmails: res.value,
            });
            console.log(res.value);
            //console.log(res);
           //console.log(err);
          });
      });
  }

  public render(): React.ReactElement<IGraphApi1Props> {

    return (
      <><div> email graph</div>
      <div>
        {this.state.AllEmails.map((email => {
          return (
            <div>
              <p className={styles.from}>{email.from.emailAddress.name}</p>
              <p>{email.subject}</p>
              <p>{email.receivedDateTime}</p>
              <p>{email.bodyPreview}</p>
            </div>
          );
        }))}
      </div></>
    );
  }
}
