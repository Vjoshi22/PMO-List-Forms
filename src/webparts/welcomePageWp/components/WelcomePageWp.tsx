import * as React from 'react';
import styles from './WelcomePageWp.module.scss';
import { IWelcomePageWpProps } from './IWelcomePageWpProps';
import { escape } from '@microsoft/sp-lodash-subset';
//external imports
import { SPComponentLoader } from "@microsoft/sp-loader";
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientConfiguration, SPHttpClientResponse, HttpClientResponse } from "@microsoft/sp-http";

let _image:string = require('./Images/CEO.png');

export default class WelcomePageWp extends React.Component<IWelcomePageWpProps, {}> {
  public render(): React.ReactElement<IWelcomePageWpProps> {
    SPComponentLoader.loadCss("https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css");
    return (
      <div> 
        {/* <article className="row single-post mt-5 no-gutters"> */}
        <img src={_image} alt="" style={{float:"right", height:"350px", width:'350px'}} className="ml-5"/>
                 
                 <p style={{textAlign:"justify"}}>
                   Lorem ipsum dolor sit amet, consectetur adipisicing elit. Nihil ad, ex eaque fuga minus reprehenderit asperiores earum incidunt. Possimus maiores dolores voluptatum enim soluta omnis debitis quam ab nemo necessitatibus.
                 Lorem ipsum dolor sit amet, consectetur adipisicing elit. Nihil ad, ex eaque fuga minus reprehenderit asperiores earum incidunt. Possimus maiores dolores voluptatum enim soluta omnis debitis quam ab nemo necessitatibus.
                 Lorem ipsum dolor sit amet, consectetur adipisicing elit. Nihil ad, ex eaque fuga minus reprehenderit asperiores earum incidunt. Possimus maiores dolores voluptatum enim soluta omnis debitis quam ab nemo necessitatibus.
                 <br/>
                 Lorem ipsum dolor sit amet, consectetur adipisicing elit. Nihil ad, ex eaque fuga minus reprehenderit asperiores earum incidunt. Possimus maiores dolores voluptatum enim soluta omnis debitis quam ab nemo necessitatibus.
                 Lorem ipsum dolor sit amet, consectetur adipisicing elit. Nihil ad, ex eaque fuga minus reprehenderit asperiores earum incidunt. Possimus maiores dolores voluptatum enim soluta omnis debitis quam ab nemo necessitatibus.
                 Lorem ipsum dolor sit amet, consectetur adipisicing elit. Nihil ad, ex eaque fuga minus reprehenderit asperiores earum incidunt. Possimus maiores dolores voluptatum enim soluta omnis debitis quam ab nemo necessitatibus.
                 Lorem ipsum dolor sit amet, consectetur adipisicing elit. Nihil ad, ex eaque fuga minus reprehenderit asperiores earum incidunt. Possimus maiores dolores voluptatum enim soluta omnis debitis quam ab nemo necessitatibus.
                 Lorem ipsum dolor sit amet, consectetur adipisicing elit. Nihil ad, ex eaque fuga minus reprehenderit asperiores earum incidunt. Possimus maiores dolores voluptatum enim soluta omnis debitis quam ab nemo necessitatibus.
           </p>
        {/* </article> */}
      </div>

    //   <div className="container">
    // {/* <article className="row single-post mt-5 no-gutters"> */}
    //     <div className="col-md-6">
    //         <div className="image-wrapper float-left pr-3">
    //             <img src="https://placeimg.com/150/150/animals" alt="">
    //         </div>
    //         <div className="single-post-content-wrapper p-3">
    //             Lorem ipsum dolor sit amet, consectetur adipisicing elit. Nihil ad, ex eaque fuga minus reprehenderit asperiores earum incidunt. Possimus maiores dolores voluptatum enim soluta omnis debitis quam ab nemo necessitatibus.
    //             <br><br>
    //             Lorem ipsum dolor sit amet, consectetur adipisicing elit. Nihil ad, ex eaque fuga minus reprehenderit asperiores earum incidunt. Possimus maiores dolores voluptatum enim soluta omnis debitis quam ab nemo necessitatibus.
    //         </div>
    //     </div>  
    //   </div>
    );
  }
}
