using DocsConverter;

var builder = WebApplication.CreateBuilder(args);
builder.Services.AddControllers();
builder.Services.AddSingleton<Hwp2Pdf>();
var app = builder.Build();
app.UseRouting();
app.MapControllers();

PrintButtonClicker clicker = new();
clicker.clickerThread.Start();

app.Run("http://0.0.0.0:5000"); // 모든 IP 주소에서 수신 대기