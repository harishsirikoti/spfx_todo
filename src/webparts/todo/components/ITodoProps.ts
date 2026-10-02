import { ITodoItem } from "../models/ITodoItem";

export interface ITodoProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  userDisplayName: string;

  // Tasks received from the SharePoint TodoTasks list.
  todoItems: ITodoItem[];
  onAddTodo: (title: string) => Promise<void>;
  onToggleTodo: (id: number, completed: boolean) => Promise<void>;
  onDeleteTodo: (id: number) => Promise<void>;
  onEditTodo: (id: number, title: string) => Promise<void>;
}
