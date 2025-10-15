using System;

namespace WinNetConfigurator.Services
{
    public class AuditService
    {
        readonly DbService _db;
        public AuditService(DbService db) { _db = db; }

        public void Log(string evt, object details = null)
        {
            _db.AddAudit(evt, details);
        }
    }
}
