using ExcelDataReader;
using Microsoft.AspNetCore.Mvc;

namespace SeatNumberSearch.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class SearchController : ControllerBase
    {

        // GET: api/TodoItems/5
        // <snippet_GetByID>
        // [HttpGet("GetStudent/{seatNumber:int}")]

        [HttpGet("GetStudent")]
        public IActionResult GetStudent([FromQuery] int seatNumber, [FromQuery] bool type)
        {

            try
            {


                var students = new List<Student>();

                var fileName = "./input/inputdata.xlsx";
                // For .net core, the next line requires the NuGet package, 
                // System.Text.Encoding.CodePages
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using (var stream = System.IO.File.Open(fileName, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {

                        while (reader.Read()) //Each row of the file
                        {

                            if (type == true && Convert.ToString(reader.GetValue(5)) == Convert.ToString(seatNumber))
                            {


                                students.Add(new Student
                                {
                                    Srno = Convert.ToString(reader.GetValue(0)),
                                    Wing = Convert.ToString(reader.GetValue(1)),
                                    Floor = Convert.ToString(reader.GetValue(2)),
                                    Block = Convert.ToString(reader.GetValue(3)),
                                    EnrollmentNo = Convert.ToString(reader.GetValue(4)),
                                    SeatNo = Convert.ToString(reader.GetValue(5)),
                                    Year = Convert.ToString(reader.GetValue(6)),
                                    Branch = Convert.ToString(reader.GetValue(7)),
                                    Name = Convert.ToString(reader.GetValue(8)),
                                    BatchNo = Convert.ToString(reader.GetValue(9)),
                                });
                                // return Ok(students);
                            }
                            else if (type == false && Convert.ToString(reader.GetValue(4)) == Convert.ToString(seatNumber))
                            {
                                students.Add(new Student
                                {
                                    Srno = Convert.ToString(reader.GetValue(0)),
                                    Wing = Convert.ToString(reader.GetValue(1)),
                                    Floor = Convert.ToString(reader.GetValue(2)),
                                    Block = Convert.ToString(reader.GetValue(3)),
                                    EnrollmentNo = Convert.ToString(reader.GetValue(4)),
                                    SeatNo = Convert.ToString(reader.GetValue(5)),
                                    Year = Convert.ToString(reader.GetValue(6)),
                                    Branch = Convert.ToString(reader.GetValue(7)),
                                    Name = Convert.ToString(reader.GetValue(8)),
                                    BatchNo = Convert.ToString(reader.GetValue(9)),
                                });

                            }


                        }
                    }
                }

                if (students == null)
                {
                    return NotFound();
                }
                else
                {
                    return Ok(students);
                }


            }
            catch (Exception exceptionMessage)
            {

                var errorObjectResult = new ObjectResult(exceptionMessage);
                errorObjectResult.StatusCode = StatusCodes.Status500InternalServerError;

                return errorObjectResult;
            }
        }

        [HttpGet("GetALLStudent")]
        public IActionResult GetALLStudent()
        {

            try
            {
                var students = new List<Student>();

                var fileName = "./input/inputdata.xlsx";
                // For .net core, the next line requires the NuGet package, 
                // System.Text.Encoding.CodePages
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using (var stream = System.IO.File.Open(fileName, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {

                        while (reader.Read()) //Each row of the file
                        {
                            if (Convert.ToString(reader.GetValue(5)) != "SeatNo")
                            {

                                students.Add(new Student
                                {
                                    Srno = Convert.ToString(reader.GetValue(0)),
                                    Wing = Convert.ToString(reader.GetValue(1)),
                                    Floor = Convert.ToString(reader.GetValue(2)),
                                    Block = Convert.ToString(reader.GetValue(3)),
                                    EnrollmentNo = Convert.ToString(reader.GetValue(4)),
                                    SeatNo = Convert.ToString(reader.GetValue(5)),
                                    Year = Convert.ToString(reader.GetValue(6)),
                                    Branch = Convert.ToString(reader.GetValue(7)),
                                    Name = Convert.ToString(reader.GetValue(8)),
                                    BatchNo = Convert.ToString(reader.GetValue(9)),
                                });
                                // return Ok(students);
                            }


                        }
                    }
                }

                if (students == null)
                {
                    return NotFound();
                }
                else
                {
                    //  var result = students.RemoveAll(s => s.SeatNo == "");
                    return Ok(students);
                }


            }
            catch (Exception exceptionMessage)
            {

                var errorObjectResult = new ObjectResult(exceptionMessage);
                errorObjectResult.StatusCode = StatusCodes.Status500InternalServerError;

                return errorObjectResult;
            }
        }
    }
}
