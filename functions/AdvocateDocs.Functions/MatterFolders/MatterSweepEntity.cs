using Microsoft.Azure.Functions.Worker;
using Microsoft.DurableTask.Entities;

namespace AdvocateDocs.Functions.MatterFolders;

/// <summary>
/// One entity per user and matter: Durable runs an entity's operations one at a time, so sweeps of the same matter
/// never overlap (two at once would file the same message twice) and a notification arriving mid-sweep gets its own
/// sweep afterwards. State = number of sweeps run.
/// </summary>
public sealed class MatterSweepEntity(MatterFolderSweeper sweeper) : TaskEntity<int>
{
    public const string SweepOperation = nameof(Sweep);
    public const string MaintainOperation = nameof(Maintain);
    private const char KeySeparator = '|';

    public static EntityInstanceId IdFor(string userId, string matterRef) => new(nameof(MatterSweepEntity), $"{userId}{KeySeparator}{matterRef}");

    /// <summary>Files what is waiting in the matter's mail folders.</summary>
    public Task Sweep() => SweepAsync(renewSubscriptions: false);

    /// <summary>Daily: renews the folders' change notifications, then sweeps as a safety net for missed ones.</summary>
    public Task Maintain() => SweepAsync(renewSubscriptions: true);

    private async Task SweepAsync(bool renewSubscriptions)
    {
        var key = Context.Id.Key.Split(KeySeparator, 2);
        await sweeper.SweepStoredAsync(key[0], key[1], renewSubscriptions, CancellationToken.None);
        State++;
    }

    [Function(nameof(MatterSweepEntity))]
    public static Task DispatchAsync([EntityTrigger] TaskEntityDispatcher dispatcher) => dispatcher.DispatchAsync<MatterSweepEntity>();
}
