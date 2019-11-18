import {IListItem} from './IListItem';

export interface ISpCrudState {
  isVisible:boolean
  status?:string;
  Items:IListItem[];
  itemTitle?:string;
  itemId?:String;
  isAscedingSort:boolean
}
