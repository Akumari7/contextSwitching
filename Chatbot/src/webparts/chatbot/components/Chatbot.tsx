import * as React from 'react';
import styles from './Chatbot.module.scss';
import { IChatbotProps } from './IChatbotProps';
// import { escape } from '@microsoft/sp-lodash-subset';
import "../utilities/webchat.js";
// import { SPServices } from './services/Services';
// import { IServices } from './services/IServices';

export default class ChatbotWebpart extends React.Component<IChatbotProps, { checked?: boolean }> {

  constructor(props: IChatbotProps) {
    
    super(props);

    this.state = {
      checked: false
    };

  }

  public render(): React.ReactElement<IChatbotProps> {
    const styleOptions = {
      // Add styleOptions to customize web chat canvas
      hideUploadButton: true
    };

    const theURL = "https://powerva.microsoft.com/api/botmanagement/v1/directline/directlinetoken?botId=" + this.props.botid;

    const store = (window as any).WebChat.createStore(
      {},
      ({ dispatch } : any) => ({next} : any) => ({ action } : any) => {
        if (action.type === "DIRECT_LINE/CONNECT_FULFILLED") {
          dispatch({
            meta: {
              method: "keyboard",
            },
            payload: {
              activity: {
                channelData: {
                  postBack: true,
                },
                //Web Chat will show the 'Greeting' System Topic message which has a trigger-phrase 'hello'
                name: 'startConversation',
                type: "event",
                value: { BotEmail: this.props.userEmail, BotNameVar: this.props.userName }
              },
            },
            type: "DIRECT_LINE/POST_ACTIVITY",
          });
        }
        return next(action);
      }
    );
    fetch(theURL)
      .then(response => response.json())
      .then(conversationInfo => {
        (window as any).WebChat.renderWebChat(
          {
            directLine: (window as any).WebChat.createDirectLine({
              token: conversationInfo.token,
            }),
            store: store,
            styleOptions: styleOptions
          },
          document.getElementById('webchat')
        );
      })
      .catch(err => console.error("An error occurred: " + err));


    /***--------------------------------******/

    // let theURL = "https://33c73de812e242e89d253734721bf6.16.environment.api.powerplatform.com/powervirtualagents/botsbyschema/kr_materialAdvisor/directline/token?api-version=2022-03-01-preview";
    // let environmentEndPoint = theURL.slice(0, theURL.indexOf('/powervirtualagents'));
    // let apiVersion = theURL.slice(theURL.indexOf('api-version')).split('=')[1];
    // let regionalChannelSettingsURL = `${environmentEndPoint}/powervirtualagents/regionalchannelsettings?api-version=${apiVersion}`;
    // let directline : any;

    // fetch(regionalChannelSettingsURL)

    //   .then((response) => {
    //     return response.json();
    //   })

    //   .then((data) => {
    //     directline = data.channelUrlsById.directline;
    //   })
    //   .catch(err => console.error("An error occurred: " + err));


    // // Triggers bot with initial message, in order to have greeting message render on load.
    // const store = (window as any).WebChat.createStore(
    //   {},
    //   ({ dispatch } : any) => ({next} : any) => ({ action } : any) => {
    //     if (action.type === "DIRECT_LINE/CONNECT_FULFILLED") {
    //       dispatch({
    //         meta: {
    //           method: "keyboard",
    //         },

    //         payload: {
    //           activity: {
    //             channelData: {
    //               postBack: true,
    //             },

    //             //Web Chat will show the 'Greeting' System Topic message which has a trigger-phrase 'hello'
    //             name: 'startConversation',
    //             type: "event",
    //             value: { BotEmail: this.props.userEmail, BotNameVar: this.props.userName }
    //           },
    //         },
    //         type: "DIRECT_LINE/POST_ACTIVITY",
    //       });
    //     }

    //     return next(action);

    //   }

    // );

    // fetch(theURL)
    //   .then(response => response.json())
    //   .then(conversationInfo => {

    //     (window as any).WebChat.renderWebChat(
    //       {
    //         directLine: (window as any).WebChat.createDirectLine({
    //           domain: `${directline}v3/directline`,
    //           token: conversationInfo.token,
    //         }),

    //         store: store,
    //         styleOptions
    //       },

    //       document.getElementById('webchat')
    //     );
    //   })

    //   .catch(err => console.error("An error occurred: " + err));

    return (

      <div className={styles.chatbotWebpart}>

        {(this.state.checked) ?
          (
            <div className={styles.container}>
              <div className={styles.row}>
                <div className={styles.header} id="header">
                  <div className={styles.header_content_container}>
                    <div className={styles.header_image_container}>
                      <img className={styles.header_image} src={this.props.botlogo} />
                    </div>
                    <div className={styles.header_title_container}>
                      <div className={styles.header_title}>
                        <span className={styles.title_text}>{this.props.botname}</span>
                        <span className={styles.close} onClick={() => { this.setState({ checked: !this.state.checked }) }}>x</span>
                      </div>
                    </div>
                  </div>
                </div>
                <div className={styles.webchat} id="webchat" role="main"></div>
              </div></div>)
          : (
            <div className={styles.botimage_container}>
              <div className={styles.chatbot_image}>
                <img src={this.props.botimage} style={{ maxHeight: 150 }} onClick={() => { this.setState({ checked: !this.state.checked }) }} />
              </div>
            </div>)
        }
      </div >
    );
  }
}
