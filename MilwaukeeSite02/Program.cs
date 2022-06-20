using MilwaukeeSite02.Models;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllersWithViews();
builder.Services.AddControllers();
builder.Services.AddRazorPages();
builder.Services.AddSingleton<List<ChangeOverModel>>();
builder.Services.AddSingleton<List<TestDataModel>>();
builder.Services.AddSingleton<List<ProcessControlModel>>();


builder.Services.AddMvc();

var app = builder.Build();
//app.Services.GetService<List<ChangeOverModel>>();


// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
}
app.UseStaticFiles();

app.UseRouting();

app.UseAuthorization();

app.MapControllers();
app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.Run();
