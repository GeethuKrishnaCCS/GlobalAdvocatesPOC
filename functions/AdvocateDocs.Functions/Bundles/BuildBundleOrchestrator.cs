using Microsoft.Azure.Functions.Worker;
using Microsoft.DurableTask;

namespace AdvocateDocs.Functions.Bundles;

/// <summary>Plan → convert members (batched fan-out) → assemble → publish. Progress is exposed as custom status.</summary>
public static class BuildBundleOrchestrator
{
    /// <summary>Members converted in parallel; keeps Graph throttling and memory in check.</summary>
    private const int ConvertBatchSize = 8;

    private static readonly TaskOptions Retry = TaskOptions.FromRetryPolicy(new RetryPolicy(3, TimeSpan.FromSeconds(10), backoffCoefficient: 2));

    [Function(nameof(BuildBundleOrchestrator))]
    public static async Task<BuildResult> Run([OrchestrationTrigger] TaskOrchestrationContext context)
    {
        try
        {
            return await BuildAsync(context);
        }
        catch (TaskFailedException ex)
        {
            // Surface the activity's own message to the caller; failure details aren't always returned by the status query.
            context.SetCustomStatus(new BuildProgress("failed", 0, 0, ex.FailureDetails.ErrorMessage));
            throw;
        }
    }

    private static async Task<BuildResult> BuildAsync(TaskOrchestrationContext context)
    {
        var input = context.GetInput<BuildInput>()!;
        var workPrefix = context.InstanceId;

        context.SetCustomStatus(new BuildProgress("loading", 0, 0));
        var plan = await context.CallActivityAsync<BundlePlan>(nameof(BundleActivities.LoadPlan), input, Retry);
        if (plan.Members.Count == 0)
        {
            throw new InvalidOperationException("The bundle has no documents.");
        }

        var converted = new List<ConvertedMember>(plan.Members.Count);
        foreach (var batch in plan.Members.Chunk(ConvertBatchSize))
        {
            context.SetCustomStatus(new BuildProgress("converting", converted.Count, plan.Members.Count));
            var tasks = batch.Select(m => context.CallActivityAsync<ConvertedMember>(nameof(BundleActivities.ConvertMember),
                new ConvertInput(workPrefix, plan.SiteId, plan.DocumentsListId, m), Retry));
            converted.AddRange(await Task.WhenAll(tasks));
        }

        context.SetCustomStatus(new BuildProgress("assembling", converted.Count, plan.Members.Count));
        var assembled = await context.CallActivityAsync<AssembleOutput>(nameof(BundleActivities.AssembleBundle),
            new AssembleInput(workPrefix, plan.Title, converted), Retry);

        context.SetCustomStatus(new BuildProgress("publishing", converted.Count, plan.Members.Count));
        return await context.CallActivityAsync<BuildResult>(nameof(BundleActivities.PublishBundle),
            new PublishInput(workPrefix, input, plan, converted, assembled), Retry);
    }
}
