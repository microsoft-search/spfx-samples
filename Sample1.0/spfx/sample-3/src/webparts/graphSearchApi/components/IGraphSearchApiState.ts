import { ISearchResult } from './ISearchResult';

export interface IGraphSearchApiState {
  results: Array<ISearchResult>;
  searchFor: string;
  externalPath: string;
  resultType: string;
  includeFiles: boolean;
  includeMessages: boolean;
  includeEvents: boolean;
}