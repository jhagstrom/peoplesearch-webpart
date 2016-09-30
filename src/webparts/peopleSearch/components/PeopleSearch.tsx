import * as React from 'react';
import {  css, Persona, PersonaSize, PersonaPresence, TextField, SearchBox, Spinner } from 'office-ui-fabric-react';
import styles from '../PeopleSearch.module.scss';
import { IPeopleSearchWebPartProps } from '../IPeopleSearchWebPartProps';
import { HttpClient } from '@microsoft/sp-client-base';
import { SearchUtils, ISearchQueryResponse, IRow } from '../../SearchUtils';
import { Utils } from '../../Utils';

export interface IPeopleSearchProps extends IPeopleSearchWebPartProps {
    siteUrl: string;
    httpClient: HttpClient;
}

export interface IPeopleSearchState {
  people: IPerson[];
  loading: boolean;
  error: string;
}
export interface IPeopleSearchResult {
  imageUrl: string;
  imageInitials: string;
  primaryText: string;
  secondaryText: string;
  tertiaryText: string;
  optionalText: string;
}

export interface IPerson {
  name: string;
  email: string;
  jobTitle: string;
  department: string;
  photoUrl: string;
  profileUrl: string;
  initials: string;
  highlight: string;

}
interface ISearchResultValue {
  Key: string;
  Value: string;
}
export default class PeopleSearch extends React.Component<IPeopleSearchProps, IPeopleSearchState> {

  constructor(props:IPeopleSearchProps, state:IPeopleSearchProps){
    super(props);
    this.state = {
      people: [] as IPerson[],
      loading: false,
      error: null
    };
  }
  public render(): JSX.Element {
    const loading: JSX.Element = this.state.loading ? <div style={{margin: '0 auto'}}><Spinner label={'Loading...'} /></div> : <div/>;
    const error: JSX.Element = this.state.error ? <div><strong>Error:</strong> {this.state.error}</div> : <div/>;
    const people: JSX.Element[] = this.state.people.map((person: IPerson, i: number) => {
      return (
        <Persona
          primaryText={person.name}
          secondaryText={person.jobTitle}
          tertiaryText={person.department}
          imageInitials={person.initials}
          imageUrl={person.photoUrl}
          size={PersonaSize.large}
          presence={PersonaPresence.none}
          onClick={() => { this.navigateTo(person.profileUrl); } }
          key={person.email} />
      )
    });
    return (
      <div>
        <p>
          <span className='ms-font-xxl'>People Search</span>
        </p>
        <SearchBox onChanged={this.handleChange.bind(this)} />
        {loading}
        {error}
        {people}
        <div style={{clear: 'both'}}/>
      </div>
);
  }

  private handleChange(newValue: string): void {
    if(newValue.length > 0){
      newValue = newValue +'*';
    }
    this.loadPeople(this.props.siteUrl, newValue);
  }
 private navigateTo(url: string): void {
    window.open(url, '_blank');
  }

  private loadPeople(siteUrl: string, query: string): void {
    this.props.httpClient.get(`${siteUrl}/_api/search/query?querytext='${query}'&sourceid='B09A7990-05EA-4AF9-81EF-EDFAB16C4E31'&selectproperties='Title,WorkEmail,JobTitle,Department,School'&hithighlightedproperties='Title,WorkEmail,JobTitle,Department,School'`, {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'odata-version': ''
      }
    })
      .then((response: Response): Promise<ISearchQueryResponse> => {
        return response.json();
      })
      .then((response: ISearchQueryResponse): void => {
        if (!response ||
          !response.PrimaryQueryResult ||
          !response.PrimaryQueryResult.RelevantResults ||
          response.PrimaryQueryResult.RelevantResults.RowCount === 0) {
          this.setState({
            loading: false,
            error: null,
            people: []
          });
          return;
        }
        const people: IPerson[] = [];
        for (let i: number = 0; i < response.PrimaryQueryResult.RelevantResults.Table.Rows.length; i++) {
          const personRow: IRow = response.PrimaryQueryResult.RelevantResults.Table.Rows[i];
          const email: string = SearchUtils.getValueFromResults('WorkEmail', personRow.Cells);
          const fullName = SearchUtils.getValueFromResults('Title', personRow.Cells);
          people.push({
            name: fullName,
            email: email,
            jobTitle: SearchUtils.getValueFromResults('JobTitle', personRow.Cells),
            department: SearchUtils.getValueFromResults('Department', personRow.Cells),
            photoUrl: Utils.getUserPhotoUrl(email, siteUrl, 'M'),
            profileUrl: SearchUtils.getValueFromResults('Path', personRow.Cells),
            initials: Utils.getInitialsFromFullName(fullName),
            highlight: SearchUtils.getValueFromResults('HitHighlightedProperties', personRow.Cells)
          });
        }
        this.setState({
          loading: false,
          error: null,
          people: people
        });
      }, (error: any): void => {
        this.setState({
          loading: false,
          error: error,
          people: []
        });
      });
  }
}
