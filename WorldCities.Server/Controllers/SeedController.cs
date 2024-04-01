using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System.Security;
using WorldCities.Server.Data;
using WorldCities.Server.Data.Models;

namespace WorldCities.Server.Controllers
{
    [Route("api/[controller]/[action]")]
    [ApiController]
    public class SeedController : ControllerBase
    {
        private readonly ApplicationDbContext _context;
        private readonly IWebHostEnvironment _env;

        public SeedController(
            ApplicationDbContext context, 
            IWebHostEnvironment env
            )
        {
            _context = context;
            _env = env;
        }

        [HttpGet]
        public async Task<ActionResult> Import()
        {
            if (!_env.IsDevelopment())
            {
                throw new SecurityException("Not allowed");
            }
            var path = Path.Combine(
                _env.ContentRootPath,
                "Data/Source/worldcities.xlsx"
                );

            using var stream = System.IO.File.OpenRead(path);
            using var excelPackage = new ExcelPackage(stream);
            // get the first worksheet
            var worksheet = excelPackage.Workbook.Worksheets[0];
            // define how many row we want to process
            var nEndRow = worksheet.Dimension.End.Row;
            var numberOfCountriesAdded = 0;
            var numberOfCitiesAdded = 0;
            // create a lookup dictionary
            // containing all the countries already existed
            // into the Database (it will be empty on first run)
            var countriesByName = _context.Countries
                .AsNoTracking()
                .ToDictionary(x => x.Name, StringComparer.OrdinalIgnoreCase);

            // iterates through all rows , skipping the first one
            for (int nRow = 2; nRow < nEndRow; nRow++)
            {
                var row = worksheet.Cells[nRow, 1, nRow, worksheet.Dimension.End.Column];
                var countryName = row[nRow, 5].GetValue<string>();
                var iso2 = row[nRow, 6].GetValue<string>();
                var iso3 = row[nRow, 7].GetValue<string>();
                // skip this country if it already exists in the database
                if (countriesByName.ContainsKey(countryName)) continue;
                var country = new Country
                {
                    Name = countryName,
                    ISO2 = iso2,
                    ISO3 = iso3,
                };
                // add new country to DB context
                await _context.Countries.AddAsync(country);
                // store the country in our lookup to retrieve its Id later on
                countriesByName.Add(countryName, country);
                // increment the counter
                numberOfCountriesAdded++;
            }

            // save all the countries into database
            if (numberOfCountriesAdded > 0)
            {
                await _context.SaveChangesAsync();
            }

            // create a lookup dictionary
            // containing all cities already existing
            // into the database (it will be empty on first run)
            var citiesByName = _context.Cities
                .AsNoTracking()
                .ToDictionary(x => (
                    Name: x.Name,
                    Lat: x.Lat,
                    Lon: x.Lon,
                    CountryId: x.CountryId
                ));

            for (int nRow = 2; nRow < nEndRow; nRow++)
            {
                var row = worksheet.Cells[
                    nRow, 1, nRow, worksheet.Dimension.End.Column
                    ];
                var name = row[nRow, 1].GetValue<string>();
                var lat = row[nRow, 3].GetValue<decimal>();
                var lon = row[nRow, 4].GetValue<decimal>();
                var countryName = row[nRow, 5].GetValue<string>();

                var countryId = countriesByName[countryName].Id;
                if (citiesByName.ContainsKey((
                    Name: name,
                    Lat: lat,
                    Lon: lon,
                    CountryId: countryId
                    ))) continue;

                var city = new City
                {
                    Name = name,
                    Lat = lat,
                    Lon = lon,
                    CountryId = countryId
                };

                _context.Cities.Add(city);
                numberOfCitiesAdded++;
            }
            if (numberOfCitiesAdded > 0)
            {
                await _context.SaveChangesAsync();
            }
            return new JsonResult(new
            {
                Cities = numberOfCitiesAdded,
                Countries = numberOfCountriesAdded
            });
        }

    }
}
