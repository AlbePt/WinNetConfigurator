using System;
using System.Collections.Generic;
using System.Linq;
using WinNetConfigurator.Models;

namespace WinNetConfigurator.Services
{
    public class IpPolicyService
    {
        readonly DbService _db;
        readonly IpPlanner _planner;
        readonly IpPolicyState _state;

        public IpPolicyState State => _state;

        public IpPolicyService(DbService db, IpPlanner planner)
        {
            _db = db ?? throw new ArgumentNullException(nameof(db));
            _planner = planner ?? throw new ArgumentNullException(nameof(planner));
            _state = BuildDefaultState();
        }

        IpPolicyState BuildDefaultState()
        {
            var state = new IpPolicyState
            {
                DefaultCabinetLimit = 8
            };

            state.CabinetPolicies["101"] = new CabinetPolicy { CabinetName = "101", Limit = 5 };
            state.CabinetPolicies["102"] = new CabinetPolicy { CabinetName = "102", Limit = 6 };
            state.CabinetPolicies["Lab-A"] = new CabinetPolicy { CabinetName = "Lab-A", Limit = 12 };

            state.Reservations.Add(new IpReservation
            {
                RangeStart = "10.0.0.1",
                RangeEnd = "10.0.0.10",
                Reason = "Резерв под инфраструктуру",
                HardBlock = true
            });
            state.Reservations.Add(new IpReservation
            {
                RangeStart = "10.0.0.200",
                RangeEnd = "10.0.0.210",
                Reason = "Гостевой сегмент",
                HardBlock = false
            });

            state.PresumedAddresses.Add(new PresumedUsage
            {
                Ip = "10.0.0.55",
                Source = "Скан сети",
                ObservedAt = DateTime.UtcNow.AddDays(-1)
            });

            return state;
        }

        public AssignmentDecision Evaluate(AppSettings settings, AssignmentDraft draft)
        {
            if (settings == null) throw new ArgumentNullException(nameof(settings));
            if (draft == null) throw new ArgumentNullException(nameof(draft));

            var decision = new AssignmentDecision();
            var busy = _db.BusyIps();

            string candidate = string.IsNullOrWhiteSpace(draft.RequestedIp)
                ? SuggestNext(settings, busy)
                : draft.RequestedIp.Trim();

            decision.SuggestedIp = candidate;

            if (busy.Contains(candidate))
            {
                decision.HasConflicts = true;
                decision.RequiresSeniorApproval = true;
                decision.Warnings.Add(new IpPolicyWarning
                {
                    Code = "busy",
                    Title = "IP занят",
                    Message = "Выбранный адрес уже присутствует в базе.",
                    RequiresSeniorApproval = true,
                    RiskFactor = RiskFactor.Conflict
                });
            }

            foreach (var reservation in _state.Reservations.Where(r => r.Contains(candidate)))
            {
                decision.IsInReservation = true;
                decision.Warnings.Add(new IpPolicyWarning
                {
                    Code = "reserve",
                    Title = reservation.HardBlock ? "Зарезервированный диапазон" : "Предупреждение",
                    Message = reservation.Reason,
                    RequiresSeniorApproval = reservation.HardBlock,
                    RiskFactor = RiskFactor.Reservation
                });
                if (reservation.HardBlock)
                    decision.RequiresSeniorApproval = true;
            }

            foreach (var presumed in _state.PresumedAddresses.Where(p => string.Equals(p.Ip, candidate, StringComparison.OrdinalIgnoreCase)))
            {
                decision.Warnings.Add(new IpPolicyWarning
                {
                    Code = "presumed",
                    Title = "Предположительно занят",
                    Message = $"Источник: {presumed.Source}, обнаружено {presumed.ObservedAt:g}",
                    RequiresSeniorApproval = presumed.RequiresConfirmation,
                    RiskFactor = RiskFactor.PresumedOccupied
                });
                if (presumed.RequiresConfirmation)
                    decision.RequiresSeniorApproval = true;
            }

            if (string.IsNullOrWhiteSpace(draft.Cabinet))
            {
                decision.Warnings.Add(new IpPolicyWarning
                {
                    Code = "cabinet_missing",
                    Title = "Кабинет",
                    Message = "Не указан кабинет.",
                    RequiresSeniorApproval = true,
                    RiskFactor = RiskFactor.UnknownCabinet
                });
                decision.RequiresSeniorApproval = true;
            }
            else
            {
                var cabinet = _db.GetCabinets().FirstOrDefault(c => string.Equals(c.Name, draft.Cabinet, StringComparison.OrdinalIgnoreCase));
                if (cabinet == null)
                {
                    decision.Warnings.Add(new IpPolicyWarning
                    {
                        Code = "cabinet_unknown",
                        Title = "Кабинет",
                        Message = "Кабинет не найден в базе.",
                        RequiresSeniorApproval = true,
                        RiskFactor = RiskFactor.UnknownCabinet
                    });
                    decision.RequiresSeniorApproval = true;
                }
                else
                {
                    var policy = _state.GetPolicyForCabinet(draft.Cabinet);
                    var limit = policy?.Limit ?? _state.DefaultCabinetLimit;
                    var count = _db.CountCabinetMachines(cabinet.Id);
                    if (count >= limit && (policy == null || !policy.Exceptions.Contains(draft.Cabinet)))
                    {
                        decision.IsCabinetLimitReached = true;
                        decision.RequiresSeniorApproval = true;
                        decision.Warnings.Add(new IpPolicyWarning
                        {
                            Code = "limit",
                            Title = "Лимит кабинета",
                            Message = $"У кабинета уже {count} адресов (лимит {limit}).",
                            RequiresSeniorApproval = true,
                            RiskFactor = RiskFactor.Limit
                        });
                    }
                }
            }

            if (draft.Attributes == null || draft.Attributes.Count == 0)
            {
                decision.Warnings.Add(new IpPolicyWarning
                {
                    Code = "attrs",
                    Title = "Атрибуты",
                    Message = "Не заполнены обязательные атрибуты устройства.",
                    RequiresSeniorApproval = false,
                    RiskFactor = RiskFactor.MissingAttributes
                });
            }

            foreach (var warning in decision.Warnings)
            {
                if (!decision.RiskFactors.Contains(warning.RiskFactor))
                {
                    decision.RiskFactors.Add(warning.RiskFactor);
                }
                if (warning.RequiresSeniorApproval)
                {
                    decision.RequiresSeniorApproval = true;
                }
            }

            draft.SuggestedIp = decision.SuggestedIp;
            draft.Warnings.Clear();
            draft.Warnings.AddRange(decision.Warnings);
            draft.RiskFactors.Clear();
            draft.RiskFactors.AddRange(decision.RiskFactors);
            draft.RequiresSeniorApproval = decision.RequiresSeniorApproval;

            return decision;
        }

        string SuggestNext(AppSettings settings, HashSet<string> busy)
        {
            if (string.IsNullOrWhiteSpace(settings.PoolStart) || string.IsNullOrWhiteSpace(settings.PoolEnd))
                throw new InvalidOperationException("Пул адресов не настроен.");

            return _planner.SuggestNextFreeIp(settings.PoolStart, settings.PoolEnd, ip =>
            {
                if (busy.Contains(ip))
                    return true;
                return _state.Reservations.Any(r => r.HardBlock && r.Contains(ip));
            });
        }
    }
}
