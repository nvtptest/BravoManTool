using Microsoft.SqlServer.Server;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace BravoManTool
{
    public class BravoManTool
    {
        // Hàm Convert Varbinary thành File
        [SqlFunction(Name = "ufn_sys_WriteFileFromVarBinary", SystemDataAccess = SystemDataAccessKind.Read, DataAccess = DataAccessKind.Read)]
        public static string ufn_sys_WriteFileFromVarBinary(string filePath, string fileNameOrigin, string fileNameCreate, byte[] fileContent, bool isZipFile)
        {
            string resultMessage = "OK";

            string tmpPath = Path.GetTempPath();

            string fileTmp = tmpPath + fileNameOrigin;

            string fileCreate = filePath + fileNameCreate;

            try
            {
                if (isZipFile)
                {
                    File.WriteAllBytes(fileTmp, fileContent);

                    using (FileStream fs = new FileStream(fileCreate, FileMode.Create))
                    using (ZipArchive arch = new ZipArchive(fs, ZipArchiveMode.Create))
                    {
                        arch.CreateEntryFromFile(fileTmp, fileNameOrigin);
                    }
                }
                else
                {
                    File.WriteAllBytes(fileCreate, fileContent);
                }
            }
            catch (Exception ex)
            {
                resultMessage = ex.Message;
            }
            finally
            {
                if (File.Exists(fileTmp))
                {
                    File.Delete(fileTmp);
                }
            }

            return resultMessage;
        }

        //Zip tất cả các file trong XML (SELECT bảng thành XML có chứa các cột: FileName, FileContent)
        [SqlProcedure]
        public static void zipFromListFileXML(SqlXml inputXml, string fileNameCreate, out string result)
        {
            result = "OK";

            string tmpPath = Path.GetTempPath();

            string fileTmp, fileName;

            List<string> listFileTmp = new List<string>();

            byte[] fileContent;

            try
            {
                XmlDocument xml = new XmlDocument();
                xml.Load(inputXml.CreateReader());

                XmlNodeReader xmlReader = new XmlNodeReader(xml);

                DataSet ds = new DataSet();
                ds.ReadXml(xmlReader);

                using (FileStream fs = new FileStream(fileNameCreate, FileMode.Create))
                using (ZipArchive arch = new ZipArchive(fs, ZipArchiveMode.Create))
                {
                    foreach (DataRow row in ds.Tables[0].Rows)
                    {
                        fileName = row["FileName"].ToString();

                        fileTmp = tmpPath + fileName;

                        fileContent = Convert.FromBase64String(row["FileContent"].ToString());

                        File.WriteAllBytes(fileTmp, fileContent);

                        listFileTmp.Add(fileTmp);

                        arch.CreateEntryFromFile(fileTmp, fileName);
                    }
                }
            }
            catch (Exception ex)
            {
                result = ex.Message;
            }
            finally
            {
                foreach (string file in listFileTmp)
                {
                    if (File.Exists(file))
                    {
                        File.Delete(file);
                    }
                }
            }
        }


        //
        [SqlProcedure]
        public static void writeFileFromVarBinary(string filePath, string fileNameOrigin, string fileNameCreate,
            byte[] fileContent, bool isZipFile, out string result)
        {
            result = "OK";

            string tmpPath = Path.GetTempPath();

            string fileTmp = tmpPath + fileNameOrigin;

            string fileCreate = filePath + fileNameCreate;

            try
            {
                if (isZipFile)
                {
                    File.WriteAllBytes(fileTmp, fileContent);

                    using (FileStream fs = new FileStream(fileCreate, FileMode.Create))
                    using (ZipArchive arch = new ZipArchive(fs, ZipArchiveMode.Create))
                    {
                        arch.CreateEntryFromFile(fileTmp, fileNameOrigin);
                    }
                }
                else
                {
                    File.WriteAllBytes(fileCreate, fileContent);
                }
            }
            catch (Exception ex)
            {
                result = ex.Message;
            }
            finally
            {
                if (File.Exists(fileTmp))
                {
                    File.Delete(fileTmp);
                }
            }
        }

        [SqlProcedure]
        public static IEnumerable readLayout(SqlXml inputXml, string fileNameCreate)
        {
            string query = "SELECT TOP 1 [id]" +
               "FROM B00UserList";

            DataTable results = new DataTable();

            using (SqlConnection conn = new System.Data.SqlClient.SqlConnection("context connection = true"))
            using (SqlCommand command = new SqlCommand(query, conn))
            using (SqlDataAdapter dataAdapter = new SqlDataAdapter(command))
                dataAdapter.Fill(results);

            DataRow dummyRow = results.NewRow();

            dummyRow[0] = "1234";
            results.Rows.Add(dummyRow);

            return results.Rows;
        }

        [SqlFunction(IsDeterministic = true, IsPrecise = true)]
        public static SqlBinary NVarCharToUtf8(SqlString inputText)
        {
            if (inputText.IsNull)
                return new SqlBinary(); // (null)

            return new SqlBinary(Encoding.UTF8.GetBytes(inputText.Value));
        }

        [SqlFunction(IsDeterministic = true, IsPrecise = true)]
        public static SqlString Utf8ToNVarChar(SqlBinary inputBytes)
        {
            if (inputBytes.IsNull)
                return new SqlString(); // (null)

            return new SqlString(Encoding.UTF8.GetString(inputBytes.Value));
        }

        [SqlFunction(IsDeterministic = true, IsPrecise = true)]
        public static SqlString RtfToText(SqlString inputText)
        {
            // Convert the RTF to plain text.
            
            if (inputText.IsNull)
                return new SqlString("");

            System.Windows.Forms.RichTextBox rtBox = new System.Windows.Forms.RichTextBox();
            try
            {
                rtBox.Rtf = inputText.Value;
            }
            catch
            {
                rtBox.Dispose();
                return new SqlString("Error load rtf");
            }

            string plainText = "";

            try
            {
                plainText = rtBox.Text;
            }
            catch
            {
                rtBox.Dispose();
                return new SqlString("Error convert rtf to text");
            }

            rtBox.Dispose();
            return new SqlString(plainText);
        }

        //Zip Folder
        [SqlProcedure]
        public static void zipFolder(string startPath, string zipPath, out string result)
        {
            result = "OK";

            try
            {
                DateTime dateStart = DateTime.Now;

                ZipFile.CreateFromDirectory(startPath, zipPath);

                DateTime dateEnd = DateTime.Now;

                result += ";" + dateStart.ToString("dd/MM/yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture) +
                           ";" + dateEnd.ToString("dd/MM/yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
            }
            catch (Exception ex)
            {
                result = ex.Message;
            }
        }


        public static void GetFileBytes(string filePath, out SqlBytes fileBytes, out SqlString resultMessage)
        {
            try
            {
                fileBytes = new SqlBytes(File.ReadAllBytes(filePath));
                resultMessage = "OK";
            }
            catch (Exception ex)
            {
                fileBytes = null;
                resultMessage = ex.Message;
            }
        }

        public static void GetBase64FromFilePath(string filePath, out SqlString base64Data, out SqlString resultMessage)
        {
            try
            {
                byte[] fileBytes = File.ReadAllBytes(filePath);
                base64Data = Convert.ToBase64String(fileBytes);
                resultMessage = "OK";
            }
            catch (Exception ex)
            {
                base64Data = null;
                resultMessage = ex.Message;
            }
        }


        private const string SecretKey = "B5003E2F66944CC1911171BD8E55D123"; 
        private const string InitializationVector = "250CB18359B549BB"; 

        /// <summary>
        /// Mã hóa BravoMan
        /// </summary>
        /// <param name="plaintext"></param>
        /// <returns></returns>
        [Microsoft.SqlServer.Server.SqlFunction]
        public static SqlString BravoManEncrypt(SqlString plaintext)
        {
            using (Aes aes = Aes.Create())
            {
                aes.Key = Encoding.UTF8.GetBytes(SecretKey);
                aes.IV = Encoding.UTF8.GetBytes(InitializationVector);

                ICryptoTransform encryptor = aes.CreateEncryptor(aes.Key, aes.IV);

                using (MemoryStream msEncrypt = new MemoryStream())
                {
                    using (CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                    {
                        using (StreamWriter swEncrypt = new StreamWriter(csEncrypt))
                        {
                            swEncrypt.Write(plaintext.Value);
                        }
                    }
                    return Convert.ToBase64String(msEncrypt.ToArray());
                }
            }
        }

        /// <summary>
        /// Chuyển file từ đường dẫn này sang đường dẫn khác
        /// </summary>
        /// <param name="sourceFilePath">Đường dẫn nguồn</param>
        /// <param name="destinationFilePath">Đường dẫn đích</param>
        /// <param name="zipFile">Có zip file không?</param>
        /// <param name="result">Kết quả trả về, nếu thành công là OK</param>
        [SqlProcedure]
        public static void MoveFileFromPath(string sourceFilePath, string destinationFilePath, bool zipFile, out string result)
        {
            try
            {
                // Kiểm tra file gốc có tồn tại không
                if (!File.Exists(sourceFilePath))
                {
                    result = "ERR: File gốc không tồn tại.";
                    return;
                }

                // Xử lý zip file nếu cần
                string tempFilePath = sourceFilePath;
                if (zipFile)
                {
                    string zipPath = Path.ChangeExtension(sourceFilePath, ".zip");
                    using (var zip = ZipFile.Open(zipPath, ZipArchiveMode.Create))
                    {
                        zip.CreateEntryFromFile(sourceFilePath, Path.GetFileName(sourceFilePath));
                    }
                    tempFilePath = zipPath;
                }

                // Chuyển file đến đường dẫn đích
                File.Copy(tempFilePath, destinationFilePath, true);

                // Đặt kết quả là OK nếu không có lỗi
                result = "OK";
            }
            catch (Exception ex)
            {
                // Ghi lại lỗi vào biến result
                result = "ERR: " + ex.Message;
            }
        }
    }
}
