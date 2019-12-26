---
title: Compose a list of items / tasks
date: 2019-12-25 03:00:00 pm
description: Learn how to add, remove and tag tasks
featured_image: '/images/blog/blowball.jpg'
---
### General info
For every task a new row is created in the task list inside the *Planning* sheet.
The column headers of the task list are fixed and should not be renamed as the data is identified by column headers.
In the following examples only relevant columns are shown to visualize the features.

### Add tasks
Let's imagine you want to plan a shopping trip and need to buy some items from the market. Your initial list looks like this:

| Task no. | Task name                        | Kanban list | Priority | tHash              |
|----------|----------------------------------|-------------|----------|--------------------|
| #000002  | Buy chia seeds                   | To do       | 3        | t50344759B92C4C0A3 |
| #000001  | Take bags                        | To do       | 2        | t5C33AE2C601539DCA |
| #000003  | Have a chat with the store owner | To do       | 1        | t35EBA085FF3A95514 |

Add a task by using the button *Add task* or hit the shortcut

*CTRL+ALT+N*

The list will then look like this:

| Task no. | Task name                        | Kanban list | Priority | tHash              |
|----------|----------------------------------|-------------|----------|--------------------|
| **#000004**  | **\<ENTER_NAME\>**           | **To do**   | **4**    | **tBA7FC881E2CA8611F** |
| #000002  | Buy chia seeds                   | To do       | 3        | t50344759B92C4C0A3 |
| #000001  | Take bags                        | To do       | 2        | t5C33AE2C601539DCA |
| #000003  | Have a chat with the store owner | To do       | 1        | t35EBA085FF3A95514 |

`Task no.` This number is generated. Do not change it

`Task name` Type in a name. The name is then copied to a special task sheet which stores the task's data (background action)

`Kanban list`: A new task is added to the list *To do* by default. It can later be set to *In progress* or *Done*

`Priority`: A new task always is assigned to the highest priority. The priority is an integer number. The task with the highest priority gets the highest number

`tHash`: Unique hash value generated to identify the task. Do not change it

### Delete tasks

Select any cell of a task's row (e.g. `To do`) to remove it by hitting 

*CTRL+ALT+R*

| Task no. | Task name                        | Kanban list | Priority | tHash              |
|----------|----------------------------------|-------------|----------|--------------------|
| ~~#000004~~ | ~~Buy strawberries~~            | `To do`     | ~~-4-~~    | ~~tBA7FC881E2CA8611F~~ |
| #000002  | Buy chia seeds                   | To do       | 3        | t50344759B92C4C0A3 |
| #000001  | Take bags                        | To do       | 2        | t5C33AE2C601539DCA |
| #000003  | Have a chat with the store owner | To do       | 1        | t35EBA085FF3A95514 |

### Group tasks by tagging

To group tasks tags can be assigned to them. Enter any tag value in one of three tag columns to apply a tag.
Firstly tags can be filtered with Excel filters. Secondly the spreadsheet highlights all tasks with the same tag if a tag is selected.

| Task no. | Task name                        | %Tag 1 | %Tag 2      | %Tag 3    |
|----------|----------------------------------|--------|-------------|-----------|
| #000004  | Buy chia seeds                   | Food products  |             | Important |
| #000003  | Buy strawberries                 | Food products  |             |           |
| #000002  | Take bags                        |        | Preparation |           |
| #000001  | Have a chat with the store owner |        |             | Important |

Selecting the tag `Important` will highlight two tasks **#000004** and **#000001**:

| Task no. | Task name                        | %Tag 1 | %Tag 2      | %Tag 3    |
|----------|----------------------------------|--------|-------------|-----------|
| **#000004**  | **Buy chia seeds**                   | **Food products**  |             | **Important** |
| #000003  | Buy strawberries                 | Food products  |             |           |
| #000002  | Take bags                        |        | Preparation |           |
| **#000001**  | **Have a chat with the store owner** |        |             | **`Important`** |

### Additional fields

`Comment` Leave a comment to the task here. Entered info will be stored in the task's data spreadsheet

`Indicator` Will be explained in a blog post later

`User time estimate` The time you think the task will take. See following blog posts	

`Total time` The total time you spent working on this task. Will be calculated for you

`Due date` The date the task should be finished.

`Contributor` Standard value is *Me*. Enter other peoples names if you want to track their work packages

`Finished on` A time this task was finished. This can be edited manually but is normally set by the spreadsheet when the corresponding task was set to *Done*
