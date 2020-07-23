import * as React from 'react';
import styles from './HomePage.module.scss';
import { IHomePageProps } from './IHomePageProps';
import { escape } from '@microsoft/sp-lodash-subset';
//external imports
import { Form, FormGroup, Button, FormControl, Row, Col, Image, Card, CardDeck } from "react-bootstrap";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientConfiguration, SPHttpClientResponse, HttpClientResponse } from "@microsoft/sp-http";

//declaring variable
let imageURL;
export let arr_tileNavigaiton:any = [];
export var arr_distinctParentVal:any =[];
var navigationListGUID = "3dfdae3b-3905-460d-8abb-06800cf4874f";

//declare state
export interface InavigationState{
  _response:string;
}

export default class HomePage extends React.Component<IHomePageProps, InavigationState> {
  constructor(props: IHomePageProps, state: InavigationState) {
    super(props);

    this.state = {
      _response: "false"
    }
  }
  public componentWillMount(){

    this.loadImages().then((items): void => {
      console.log(items);
      //clearing the array
      arr_tileNavigaiton.splice(0, arr_tileNavigaiton.length);
      items.forEach(item => {
        arr_tileNavigaiton.push({
          Title: item.Title,
          Parent: item.Parent,
          BackgroundImageURL: item.BackgroundImageLocation.Url,
          LinkName: item.LinkLocation.Description,
          LinkURL: item.LinkLocation.Url
        })
      });
    }).then((img) => {
      const distinctArray = (arr_tileNavigaiton.map(p => p.Parent).filter((Parent, index, arr) => arr.indexOf(Parent) == index));
      //clearing the array
      arr_distinctParentVal.splice(0,arr_distinctParentVal.length); 
      distinctArray.forEach(item => {
        arr_distinctParentVal.push(item);
        this.setState({
          _response:"true"
        })
      });
  });
  }
  public render(): React.ReactElement<IHomePageProps> {
    SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.1.3/css/bootstrap.css");
    if(this.state._response == "false"){
      return (
        <div id="maindiv">...Loading</div>
      )
    }else{
    return (
      
      <div id="maindiv">
        <CardDeck>

          {arr_tileNavigaiton.filter(homeTiles => homeTiles.Parent === this.props.tileName).map((item) =>
            <div className="mb-3">
              <Card style={{ width: '16.5rem'}}>
                <Card.Link href={item.LinkURL}>
                  <Card.Img style={{maxHeight:'170px', height:'200px'}} variant="top" src={item.BackgroundImageURL} />
                </Card.Link>
                <Card.Body className={styles.cardBodyBox}>
                  <Card.Title className={styles.navTitle} data-toggle="tooltip" data-placement="top" title={item.Title}>
                  <Card.Link className={styles.tileTitle} href={item.LinkURL} >{item.Title}</Card.Link></Card.Title>
                </Card.Body>
              </Card>
            </div>
          )
          }
        </CardDeck>
      </div>
    );}
  }
  private loadImages(): Promise<any> {
    const url = `${this.props.currentContext.pageContext.web.absoluteUrl}/_api/web/lists('${navigationListGUID}')/items?$orderby=TileOrder asc`;

    return this.props.currentContext.spHttpClient.get(url,
      SPHttpClient.configurations.v1)
      .then(response => {
        return response.json();
      }).then(jsonresponse => {
        return jsonresponse.value;
        console.log(jsonresponse.value);
      })
  }
}
