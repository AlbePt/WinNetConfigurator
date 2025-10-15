using System;
using System.Threading;

namespace WinNetConfigurator.Models
{
    public enum BackgroundTaskStatus
    {
        Pending,
        Running,
        Completed,
        Failed,
        Cancelled
    }

    public class BackgroundTask
    {
        public Guid Id { get; } = Guid.NewGuid();
        public string Name { get; }
        public BackgroundTaskStatus Status { get; private set; } = BackgroundTaskStatus.Pending;
        public int Progress { get; private set; }
        public string StateMessage { get; private set; }
        public DateTime CreatedAt { get; } = DateTime.UtcNow;
        public DateTime? CompletedAt { get; private set; }
        public string ResultMessage { get; private set; }
        public Exception Error { get; private set; }
        public CancellationTokenSource Cancellation { get; } = new CancellationTokenSource();

        public BackgroundTask(string name)
        {
            Name = string.IsNullOrWhiteSpace(name) ? "Операция" : name.Trim();
        }

        public void SetRunning()
        {
            Status = BackgroundTaskStatus.Running;
            Progress = 0;
            StateMessage = "Выполняется";
        }

        public void UpdateProgress(int percent, string message)
        {
            if (percent < 0) percent = 0;
            if (percent > 100) percent = 100;
            Progress = percent;
            StateMessage = message;
        }

        public void Complete(string resultMessage)
        {
            Status = BackgroundTaskStatus.Completed;
            Progress = 100;
            ResultMessage = resultMessage;
            CompletedAt = DateTime.UtcNow;
            StateMessage = resultMessage;
        }

        public void Fail(Exception ex)
        {
            Status = BackgroundTaskStatus.Failed;
            Error = ex;
            CompletedAt = DateTime.UtcNow;
            StateMessage = ex.Message;
        }

        public void Cancel()
        {
            if (!Cancellation.IsCancellationRequested)
            {
                Cancellation.Cancel();
            }
            Status = BackgroundTaskStatus.Cancelled;
            CompletedAt = DateTime.UtcNow;
            StateMessage = "Отменено";
        }
    }
}
