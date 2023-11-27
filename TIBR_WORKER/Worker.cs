
namespace TIBR_WORKER
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;
        private readonly TIBR_Service _tibrService;

        public Worker(ILogger<Worker> logger, TIBR_Service tibrService)
        {
            _logger = logger;
            _tibrService = tibrService;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    _tibrService.ProcessTIBRData();
                    _logger.LogInformation("TIBR data processing completed at: {time}", DateTimeOffset.Now);

                } catch (Exception ex)
                {
                    _logger.LogError(ex, "Error occurred during TIBR data processing: {message}", ex.Message);
                }
                await Task.Delay(1000 * 30, stoppingToken);
            }
        }
    }
}