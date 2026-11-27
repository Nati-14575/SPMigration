using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
// AllowTrailingCommas: arrays pasted from the IAM dump often end with a trailing comma.
builder.Services.AddControllers()
    .AddJsonOptions(o => o.JsonSerializerOptions.AllowTrailingCommas = true);
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();


// register your custom services
builder.Services.AddSingleton<SharePointService>();
builder.Services.AddSingleton<MinIOService>();
builder.Services.AddHttpClient<ApiService>();
builder.Services.AddHttpClient<AdExplorerService>();
builder.Services.AddSingleton<MigrationService>();
builder.Services.AddTransient<IamGroupMigrationService>();
builder.Services.AddTransient<IamGroupMembersService>();
builder.Services.AddTransient<IamPositionResolver>();
builder.Services.AddTransient<IamGroupSyncService>();

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();
app.UseAuthorization();
app.MapControllers();

app.Run();
