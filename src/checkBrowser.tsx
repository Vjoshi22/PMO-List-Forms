import * as React from 'react';
import { escape } from '@microsoft/sp-lodash-subset';

export default class CheckBrowser extends React.Component<{}, {}> {
  public render(): React.ReactElement<{}> {
    return (
        <div style={{backgroundColor:'yellow'}}>
            <h5>This browser version is not compitable with our website, please use latest version of the browsers</h5>
        </div>
    );
  }
}
