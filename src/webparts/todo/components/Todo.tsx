import * as React from "react";

import styles from "./Todo.module.scss";
import { ITodoProps } from "./ITodoProps";

const Todo: React.FC<ITodoProps> = (props) => {
  const [taskTitle, setTaskTitle] = React.useState<string>("");
  const [isSaving, setIsSaving] = React.useState<boolean>(false);
  const [editingId, setEditingId] = React.useState<number | null>(null);

  const [editingTitle, setEditingTitle] = React.useState<string>("");
  const handleAddTodo = async (): Promise<void> => {
    const trimmedTitle = taskTitle.trim();

    if (!trimmedTitle) {
      return;
    }

    try {
      setIsSaving(true);

      await props.onAddTodo(trimmedTitle);

      // Clear input only after SharePoint successfully saves.
      setTaskTitle("");
    } catch (error) {
      console.error("Failed to add todo:", error);
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <section className={styles.todo}>
      <h2>My Todo Tasks</h2>

      <p>Welcome, {props.userDisplayName}</p>

      <div className={styles.addTask}>
        <input
          type="text"
          value={taskTitle}
          placeholder="Enter a task"
          onChange={(event) => setTaskTitle(event.target.value)}
        />

        <button type="button" onClick={handleAddTodo} disabled={isSaving}>
          {isSaving ? "Adding..." : "Add"}
        </button>
      </div>

      <ul>
       {props.todoItems.map((item) => (

  <li key={item.Id}>

    <input
      type="checkbox"
      checked={item.Completed}
      onChange={async (event) => {
        await props.onToggleTodo(
          item.Id,
          event.target.checked
        );
      }}
    />

    {editingId === item.Id ? (

      <>
        <input
          type="text"
          value={editingTitle}
          onChange={(event) => {
            setEditingTitle(event.target.value);
          }}
        />

        <button
          type="button"
          onClick={async () => {

            const trimmedTitle =
              editingTitle.trim();

            if (!trimmedTitle) {
              return;
            }

            await props.onEditTodo(
              item.Id,
              trimmedTitle
            );

            setEditingId(null);
            setEditingTitle('');
          }}
        >
          Save
        </button>

        <button
          type="button"
          onClick={() => {
            setEditingId(null);
            setEditingTitle('');
          }}
        >
          Cancel
        </button>
      </>

    ) : (

      <>
        <span>
          {item.Title}
        </span>

        {' - '}

        <span>
          {item.Completed
            ? 'Completed'
            : 'Pending'}
        </span>

        <button
          type="button"
          onClick={() => {

            setEditingId(item.Id);
            setEditingTitle(item.Title);

          }}
        >
          Edit
        </button>

        <button
          type="button"
          onClick={async () => {
            await props.onDeleteTodo(item.Id);
          }}
        >
          Delete
        </button>
      </>

    )}

  </li>

))}
      </ul>
    </section>
  );
};

export default Todo;
