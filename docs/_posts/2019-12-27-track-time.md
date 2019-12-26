---
title: How to track the work you're doing
date: 2019-12-27 00:00:00 pm
description: Waste time on tasks
featured_image: '/images/blog/woman.jpg'
---

When adding a task you might have noticed that the spreadsheet asked you to *\<ENTER_ESTIMATE\>* in the *User time estimate* column. To improve your estimates it is helpful to enter the time you think the task needs to be done there and compare it against the actual time you spent on it after it has been finished.

To set your estimate just overwrite the placeholder with a time value in h like *5.00*

### Track your time

Tracked time will be a difference of two timestamps marking the start and end of work. To set timestamps select any cell in a task's row and use one of these buttons:

`Start task`: Will set a start timestamp to the selected task. An indicator will be set in the column *Indicator* to show which task is currently tracked.

`End task`: Will set an end timestamp to the currently tracked task. This ends tracking and the indicator will be removed. Time difference will be calculated automatically in the background

`Add 15m`: Will add 15 minutes of time to the task. the end timestamp for the current task will be set to *NOW* and the start timestamp will be adjusted.

`Add 0m`: Adds a point of time (start timestamp equals end timestamp) to a task

| Task no. | Task name                        | Indicator | User time estimate | Total time |
|----------|----------------------------------|-----------|--------------------|------------|
| #000005  | Take bags                        |           | \<ENTER_ESTIMATE\> |            |
| **#000002**  | **Buy chia seeds**                   | **\<current** | **0.25**               |            |
| #000003  | Buy strawberries                 |           | 0.50               |            |
| #000001  | Buy soy sauce                    |           | 1.50               |            |
| #000004  | Have a chat with the store owner |           | 6.00               |            |

The timestamps will be saved in the task's data sheet. Situation after `Start task` button was clicked:

| Time entry no. | Start time       | End time         | Time delta in h | Indicator |
|----------------|------------------|------------------|-----------------|-----------|
| #000001        | 18/12/2019 01:33 am | 18/12/2019 02:05 am | 0,53            |           |
| #000002        | 26/12/2019 04:10 pm | 26/12/2019 04:25 pm | 0,25            |           |
| **#000003**        | **26/12/2019 04:25 pm** |                  | **0,00**            | **<current**  |

If you made a mistake tracking a task just click on its name and edit the timestamps inside the loaded data sheet.
When I miss starting the tracking I often hit the `Add 15m` button several times to add time to a task afterwards. 

To see how much time you spent on  your task click the `Collect total times` button and wait for the script updating the column info:

| Task no. | Task name                        | Indicator | User time estimate | Total time |
|----------|----------------------------------|-----------|--------------------|------------|
| #000005  | Take bags                        |           | \<ENTER_ESTIMATE\> |  0.00      |
| #000002  | Buy chia seeds                   |           | 0.25               |  0.53      |
| #000003  | Buy strawberries                 |           | 0.50               |  0.50      |
| #000001  | Buy soy sauce                    |           | 1.50               |  0.00      |
| #000004  | Have a chat with the store owner |           | 6.00               |  0.00      |

Now you can see actual task time and your estimate in comparison.

### Additional notes on tracking

Tracking consumes time and micromanaging tasks is exhaustive. So keep tracking simple:

1. Only add as few tasks as possible to the list. A task less than 15 minutes might not be worth tracking. In the table above no estimate was entered for the task *Take bags* as it does not take very long. In conjunction no time was tracked for the task.

1. Try to stick to one task as long as you finish it. Switching tasks will increase micromanagement and is time consuming. Furthermore it is not a good idea because you need some time to dive in and out of a work package.

1. Avoid tracking two closely related tasks: If set tasks are *Run 5k* and *Drink water during run* you end up changing the currently tracked task a lot. Think about joining those two tasks and add more info to the *Comment* field of the task.

1. Instead of entering an estimate in hours you could also enter point-based estimates like Scrum story points. As long as your estimates are consistent and comparable to each other EBS algorithm will still produce valid results.

1. If a team-mate wants to give you insight into some cat videos she or he watched yesterday do not stop the tracking: Distractions are normal and part of the process. The assumption is that they happen regularily and will contribute to a difference between estimation and actual task time. EBS algorithm will cover this.

1. Later we'll discuss tracking time by using your calendar to make tracking of day-long tasks easier. Stay tuned.