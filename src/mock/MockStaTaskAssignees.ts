export interface ITaskAssigneeQuery {
  TaskAssignee:{ID:number};
  Title:string;
}

export const MockStaTaskAssignees:ITaskAssigneeQuery[] = [
    {TaskAssignee: {ID: 6}, Title: 'CEO'},
    {TaskAssignee: {ID: 5}, Title: 'Corporate Head'},
    {TaskAssignee: {ID: 3}, Title: 'Commercial Head'},
    {TaskAssignee: {ID: 7}, Title: 'Financial Director'},
  ];