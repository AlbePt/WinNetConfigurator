using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using WinNetConfigurator.Models;

namespace WinNetConfigurator.Services
{
    public class TaskQueueService
    {
        readonly List<BackgroundTask> _tasks = new List<BackgroundTask>();
        readonly object _lock = new object();

        public event EventHandler TasksChanged;

        public BackgroundTask Enqueue(string name, Func<CancellationToken, IProgress<TaskProgressReport>, Task<string>> work)
        {
            if (work == null) throw new ArgumentNullException(nameof(work));
            var task = new BackgroundTask(name);
            lock (_lock)
            {
                _tasks.Add(task);
            }
            TasksChanged?.Invoke(this, EventArgs.Empty);

            Task.Run(async () =>
            {
                var progress = new Progress<TaskProgressReport>(report =>
                {
                    if (report == null) return;
                    lock (_lock)
                    {
                        task.UpdateProgress(report.Percent, report.Message);
                    }
                    TasksChanged?.Invoke(this, EventArgs.Empty);
                });

                try
                {
                    task.SetRunning();
                    TasksChanged?.Invoke(this, EventArgs.Empty);
                    var result = await work(task.Cancellation.Token, progress).ConfigureAwait(false);
                    lock (_lock)
                    {
                        task.Complete(result);
                    }
                }
                catch (OperationCanceledException)
                {
                    lock (_lock)
                    {
                        task.Cancel();
                    }
                }
                catch (Exception ex)
                {
                    lock (_lock)
                    {
                        task.Fail(ex);
                    }
                }
                finally
                {
                    TasksChanged?.Invoke(this, EventArgs.Empty);
                }
            });

            return task;
        }

        public IReadOnlyList<BackgroundTask> ListTasks()
        {
            lock (_lock)
            {
                return _tasks.Select(t => t).ToList();
            }
        }

        public void Cancel(Guid id)
        {
            lock (_lock)
            {
                var task = _tasks.FirstOrDefault(t => t.Id == id);
                task?.Cancel();
            }
            TasksChanged?.Invoke(this, EventArgs.Empty);
        }
    }
}
