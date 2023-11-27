using TIBR_WORKER;

IHost host = Host.CreateDefaultBuilder(args)
    .ConfigureServices(services =>
    {
        services.AddTransient<TIBR_Service>();
        services.AddHostedService<Worker>();
    })
    .Build();

await host.RunAsync();
