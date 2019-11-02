import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ClientMode } from './ClientMode';

export interface IGraphSearchApiProps {
  description: string;
  clientMode: ClientMode;
  context: WebPartContext;
}
