using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.IO;

namespace cfire
{
    class Program
    {
        const string ConnStrFMT = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES"";";
        static List<dynamic> ParseHarborExcel(string filename, string sheetName)
        {
            List<dynamic> prjs = new List<dynamic>();
            if (true == File.Exists(filename))
            {
                var connStr = string.Format(ConnStrFMT, filename);
                try
                {
                    using (var conn = new OleDbConnection(connStr))
                    {
                        conn.Open();
                        using (var cmd = new OleDbCommand(string.Format("select * from [{0}$]", sheetName), conn))
                        {
                            using (var reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                    prjs.Add(new { X = reader.GetDouble(0), 
                                                   Y = reader.GetDouble(1),
                                                   Name = reader.GetString(2).Replace("\"", "'").Replace("\n", " "),
                                                   Basin = reader.GetString(3).Replace("\"", "'").Replace("\n", " "),
                                                   State = reader.GetString(4).Replace("\"", "'").Replace("\n", " "),
                                                   Contact = reader.GetString(5).Replace("\"", "'").Replace("\n", " "),
                                                   Title = reader.GetString(6).Replace("\"", "'").Replace("\n", " "),
                                                   Phone = reader.GetString(7).Replace("\"", "'").Replace("\n", " "),
                                                   Email = reader.GetString(9).Replace("\"", "'").Replace("\n", " "),
                                                   OfficeLocation = reader.GetValue(10).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   Street = reader.GetValue(11).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   CityStateZip = reader.GetValue(12).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   ContactType = reader.GetValue(13).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   AlternativeContact = reader.GetValue(14).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   AlternativePhone = reader.GetValue(15).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   AlternativeEmail = reader.GetValue(16).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   DredgingFreq = reader.GetDouble(17),
                                                   DredgingPeriod = reader.GetValue(18).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   DredgingVolume = reader.GetDouble(19),
                                                   NearshorePlacement = reader.GetDouble(20),
                                                   OpenLakePlacement = string.IsNullOrWhiteSpace(reader.GetValue(21).ToString()) ? 0.0 : double.Parse(reader.GetValue(21).ToString()),
                                                   CDFPlacement = string.IsNullOrWhiteSpace(reader.GetValue(22).ToString()) ? 0.0 : double.Parse(reader.GetValue(22).ToString()),
                                                   FacilityName = reader.GetValue(23).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   OtherBeneficialUse = reader.GetValue(24).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   OtherBeneficialUseAmount = string.IsNullOrWhiteSpace(reader.GetValue(25).ToString()) ? 0.0 : double.Parse(reader.GetValue(25).ToString()),
                                                   OtherBeneficialUseLocation = reader.GetValue(26).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   DredgingProjParam = reader.GetValue(27).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   DredgingProjDepth = reader.GetDouble(28),
                                                   DredgingProjStatus = reader.GetValue(29).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   Url = reader.GetValue(30).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   SedimentCharacter = reader.GetValue(31).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   SedimentComment = reader.GetValue(32).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   ContaminantTrend = reader.GetValue(33).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   ContaminantConcern = reader.GetValue(34).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   Comment = reader.GetValue(35).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   Metal = reader.GetValue(36).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   MetalComment = reader.GetValue(37).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   PCB = reader.GetValue(38).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   PCBComment = reader.GetValue(39).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   Pesticides = reader.GetValue(40).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   PesticidesComment = reader.GetValue(41).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   PAHs = reader.GetValue(42).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   PAHsComment = reader.GetValue(43).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   VOCs = reader.GetValue(44).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   VOCsComment = reader.GetValue(45).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   Others = reader.GetValue(46).ToString().Replace("\"", "'").Replace("\n", " "),
                                                   OthersComment = reader.GetValue(47).ToString().Replace("\"", "'").Replace("\n", " ")
                                                    });
                            }
                        }

                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
            return prjs;
        }
        static List<dynamic> ParseCDFExcel(string filename, string sheetName)
        {
            List<dynamic> prjs = new List<dynamic>();
            if (true == File.Exists(filename))
            {
                var connStr = string.Format(ConnStrFMT, filename);
                try
                {
                    using (var conn = new OleDbConnection(connStr))
                    {
                        conn.Open();
                        using (var cmd = new OleDbCommand(string.Format("select * from [{0}$]", sheetName), conn))
                        {
                            using (var reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                    prjs.Add(new
                                    {
                                        X = reader.GetDouble(0),
                                        Y = reader.GetDouble(1),
                                        Name = reader.GetString(2).Replace("\"", "'").Replace("\n", " "),
                                        Basin = reader.GetString(3).Replace("\"", "'").Replace("\n", " "),
                                        State = reader.GetString(4).Replace("\"", "'").Replace("\n", " "),
                                        ContactType = reader.GetValue(5).ToString().Replace("\"", "'").Replace("\n", " "),
                                        Contact = reader.GetString(6).Replace("\"", "'").Replace("\n", " "),
                                        Title = reader.GetString(7).Replace("\"", "'").Replace("\n", " "),
                                        Phone = reader.GetString(8).Replace("\"", "'").Replace("\n", " "),
                                        MobilePhone = reader.GetValue(9).ToString().Replace("\"", "'").Replace("\n", " "),
                                        Email = reader.GetString(10).Replace("\"", "'").Replace("\n", " "),
                                        Street = reader.GetValue(11).ToString().Replace("\"", "'").Replace("\n", " "),
                                        CityStateZip = reader.GetValue(12).ToString().Replace("\"", "'").Replace("\n", " "),
                                        CDFStatus = reader.GetValue(13).ToString().Replace("\"", "'").Replace("\n", " "),
                                        RemainingCapacity = reader.GetDouble(14),
                                        AverageMaterialRecv = reader.GetDouble(15),
                                        EstVolume = reader.GetDouble(16),
                                        CDFOwner = reader.GetValue(17).ToString().Replace("\"", "'").Replace("\n", " "),
                                        CDFOperator = reader.GetValue(18).ToString().Replace("\"", "'").Replace("\n", " "),
                                        Authority = reader.GetValue(19).ToString().Replace("\"", "'").Replace("\n", " "),
                                        StagingArea = reader.GetValue(20).ToString().Replace("\"", "'").Replace("\n", " "),
                                        RoadAccess = reader.GetValue(21).ToString().Replace("\"", "'").Replace("\n", " "),
                                        RoadAccessDetail = reader.GetValue(22).ToString().Replace("\"", "'").Replace("\n", " "),
                                        RailAccess = reader.GetValue(23).ToString().Replace("\"", "'").Replace("\n", " "),
                                        RailAccessDetail = reader.GetValue(24).ToString().Replace("\"", "'").Replace("\n", " "),
                                        WaterAccess = reader.GetValue(25).ToString().Replace("\"", "'").Replace("\n", " "),
                                        WaterAccessDetail = reader.GetValue(26).ToString().Replace("\"", "'").Replace("\n", " "),
                                        SedimentCharacter = reader.GetValue(27).ToString().Replace("\"", "'").Replace("\n", " "),
                                        NearShoreNourishment = reader.GetValue(28).ToString().Replace("\"", "'").Replace("\n", " "),
                                        ContaminantConcern = reader.GetValue(29).ToString().Replace("\"", "'").Replace("\n", " "),
                                        Comment = reader.GetValue(30).ToString().Replace("\"", "'").Replace("\n", " "),
                                        Metal = reader.GetValue(31).ToString().Replace("\"", "'").Replace("\n", " "),
                                        MetalComment = reader.GetValue(32).ToString().Replace("\"", "'").Replace("\n", " "),
                                        PCB = reader.GetValue(33).ToString().Replace("\"", "'").Replace("\n", " "),
                                        PCBComment = reader.GetValue(34).ToString().Replace("\"", "'").Replace("\n", " "),
                                        Pesticides = reader.GetValue(35).ToString().Replace("\"", "'").Replace("\n", " "),
                                        PesticidesComment = reader.GetValue(36).ToString().Replace("\"", "'").Replace("\n", " "),
                                        PAHs = reader.GetValue(37).ToString().Replace("\"", "'").Replace("\n", " "),
                                        PAHsComment = reader.GetValue(38).ToString().Replace("\"", "'").Replace("\n", " "),
                                        VOCs = reader.GetValue(39).ToString().Replace("\"", "'").Replace("\n", " "),
                                        VOCsComment = reader.GetValue(40).ToString().Replace("\"", "'").Replace("\n", " "),
                                        Others = reader.GetValue(41).ToString().Replace("\"", "'").Replace("\n", " "),
                                        OthersComment = reader.GetValue(42).ToString().Replace("\"", "'").Replace("\n", " ")
                                    });
                            }
                        }

                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
            return prjs;
        }
        static void Main(string[] args)
        {
            const string HarborJsonFmt = @"{{""lon"":{0},""lat"":{1},""name"":""{2}"",""basin"":""{3}"",""state"":""{4}"",""contact"":""{5}"",""title"":""{6}"",
                                    ""phone"":""{7}"",""email"":""{8}"",""office"":""{9}"",""street"":""{10}"",""citystatezip"":""{11}"",
                                    ""contactType"":""{12}"",""alterContact"":""{13}"",""alterPhone"":""{14}"",""alterEmail"":""{15}"",""dredgingFreq"":{16},
                                    ""dredgingPeriod"":""{17}"",""dredgingVolume"":{18},""nearshorePlace"":{19},""openLakePlace"":{20},
                                    ""CDFPlace"":{21},""facilityName"":""{22}"",""otherUse"":""{23}"",""otherUseAmnt"":{24},""OULoc"":""{25}"",
                                    ""projectParam"":""{26}"",""projDepth"":{27},""porjStatus"":""{28}"",""url"":""{29}"",""sedCharacter"":""{30}"",
                                    ""sedComment"":""{31}"",""contTrend"":""{32}"",""contConcern"":""{33}"",""contComment"":""{34}"",""metal"":""{35}"",
                                    ""metalComment"":""{36}"",""pcbs"":""{37}"",""pcbsComment"":""{38}"",""pesticides"":""{39}"",""pesticidesComment"":""{40}"",
                                    ""pahs"":""{41}"",""pahsComment"":""{42}"",""vocs"":""{43}"",""vocsComment"":""{44}"",""others"":""{45}"",""othersComment"":""{46}""}}";
            List<dynamic> list = ParseHarborExcel(@"C:\Users\gwang.GLC\Documents\Visual Studio 2013\Projects\cfire\CFIRE_160419\Harbors_CFIRE_160419.xlsx",
                                            "Harbors_USACE_2015");
            StringBuilder sb = new StringBuilder();
            sb.Append('[');
            foreach(var harbor in list)
            {
                sb.AppendFormat(HarborJsonFmt, harbor.X, harbor.Y, harbor.Name, harbor.Basin, harbor.State, harbor.Contact, harbor.Title, harbor.Phone,
                                harbor.Email, harbor.OfficeLocation, harbor.Street, harbor.CityStateZip, harbor.ContactType, harbor.AlternativeContact,
                                harbor.AlternativePhone, harbor.AlternativeEmail, harbor.DredgingFreq, harbor.DredgingPeriod, harbor.DredgingVolume,
                                harbor.NearshorePlacement, harbor.OpenLakePlacement, harbor.CDFPlacement, harbor.FacilityName, harbor.OtherBeneficialUse,
                                harbor.OtherBeneficialUseAmount, harbor.OtherBeneficialUseLocation, harbor.DredgingProjParam, harbor.DredgingProjDepth,
                                harbor.DredgingProjStatus, harbor.Url, harbor.SedimentCharacter, harbor.SedimentComment, harbor.ContaminantTrend,
                                harbor.ContaminantConcern, harbor.Comment, harbor.Metal, harbor.MetalComment, harbor.PCB, harbor.PCBComment,
                                harbor.Pesticides, harbor.PesticidesComment, harbor.PAHs, harbor.PAHsComment, harbor.VOCs, harbor.VOCsComment,
                                harbor.Others, harbor.OthersComment);
                sb.Append(',');
            }
            sb.Remove(sb.Length - 1, 1);
            sb.Append(']');
            System.IO.File.WriteAllText("usaceharbor.json", sb.ToString());


            const string CDFJsonFmt = @"{{""lon"":{0},""lat"":{1},""name"":""{2}"",""basin"":""{3}"",""state"":""{4}"",""contact"":""{5}"",""title"":""{6}"",
                                    ""phone"":""{7}"",""email"":""{8}"",""mobile"":""{9}"",""street"":""{10}"",""citystatezip"":""{11}"",
                                    ""contactType"":""{12}"",""status"":""{13}"",""remainingCapacity"":{14},""avMaterialRecv"":{15},""estVolume"":{16},
                                    ""owner"":""{17}"",""operator"":""{18}"",""authority"":""{19}"",""stagingArea"":""{20}"",
                                    ""roadAccess"":""{21}"",""roadDetail"":""{22}"",""railAccess"":""{23}"",""railDetail"":""{24}"",""waterAccess"":""{25}"",
                                    ""waterDetail"":""{26}"",""sedCharacter"":""{27}"",""nearshoreNourishment"":""{28}"",""contConcern"":""{29}"",
                                    ""contComment"":""{30}"",""metal"":""{31}"",""metalComment"":""{32}"",""pcbs"":""{33}"",""pcbsComment"":""{34}"",
                                    ""pesticides"":""{35}"",""pesticidesComment"":""{36}"",""pahs"":""{37}"",""pahsComment"":""{38}"",""vocs"":""{39}"",
                                    ""vocsComment"":""{40}"",""others"":""{41}"",""othersComment"":""{42}""}}";
            list = ParseCDFExcel(@"C:\Users\gwang.GLC\Documents\Visual Studio 2013\Projects\cfire\CFIRE_160419\CDFs_CFIRE_160419.xlsx",
                                 "CDFs_USACE_2015");
            sb.Clear();
            sb.Append('[');
            foreach(var cdf in list)
            {
                sb.AppendFormat(CDFJsonFmt,cdf.X, cdf.Y, cdf.Name, cdf.Basin, cdf.State, cdf.Contact, cdf.Title, cdf.Phone, cdf.Email, cdf.MobilePhone,
                                cdf.Street, cdf.CityStateZip, cdf.ContactType, cdf.CDFStatus, cdf.RemainingCapacity, cdf.AverageMaterialRecv, cdf.EstVolume,
                                cdf.CDFOwner, cdf.CDFOperator, cdf.Authority, cdf.StagingArea, cdf.RoadAccess, cdf.RoadAccessDetail, cdf.RailAccess,
                                cdf.RailAccessDetail, cdf.WaterAccess, cdf.WaterAccessDetail, cdf.SedimentCharacter, cdf.NearShoreNourishment,
                                cdf.ContaminantConcern, cdf.Comment, cdf.Metal, cdf.MetalComment, cdf.PCB, cdf.PCBComment,
                                cdf.Pesticides, cdf.PesticidesComment, cdf.PAHs, cdf.PAHsComment, cdf.VOCs, cdf.VOCsComment,
                                cdf.Others, cdf.OthersComment);
                sb.Append(',');
            }
            sb.Remove(sb.Length - 1, 1);
            sb.Append(']');
            System.IO.File.WriteAllText("cdf.json", sb.ToString());
        }
    }
}
