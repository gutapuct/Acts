using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Acts
{
    public class ReasonsModel
    {
        public string[] NamesOfEquipment { get; set; }
        public List<Reason> Reasons { get; set; }
    }

    public class Reason
    {
        public Reason()
        {
            this.NameReason = "Not found!!!";
            this.NameRecommendation = "Not found!!!";
        }
        public string NameReason { get; set; }
        public string NameRecommendation { get; set; }
        public bool WasUsed { get; set; }
    }

    public static class Reasons
    {
        public static List<ReasonsModel> ToReasonModel (this string[,] data)
        {
            var result = new List<ReasonsModel>();

            for (var line = 0; line < data.GetLength(1); line++)
            {
                var reasons = new List<Reason>();
                for (var column = 1; column < data.GetLength(0); column++)
                {
                    var reason = data[column, line];
                    if (!String.IsNullOrWhiteSpace(reason))
                    {
                        var reasonAndRecommendation = reason.Split('|');
                        if (reasonAndRecommendation.Where(i => i.Length > 0).Count() == 2)
                        {
                            reasons.Add(new Reason
                            {
                                NameReason = reasonAndRecommendation[0].Trim(),
                                NameRecommendation = reasonAndRecommendation[1].Trim(),
                                WasUsed = false
                            });
                        }
                    }
                }
                if (reasons.Count > 0)
                {
                    result.Add(new ReasonsModel
                    {
                        NamesOfEquipment = data[0, line].Split('|').Select(i => i = i.ToLower()).ToArray(),
                        Reasons = reasons
                    });
                }
            }

            return result;
        }

        public static Reason GetReasonByEquipmentName(this List<ReasonsModel> reasons, string equipmentName)
        {
            equipmentName = equipmentName.ToLower().Trim();

            ReasonsModel line = null;
            foreach (var reasonsModel in reasons)
            {
                var reasonModel = reasonsModel.NamesOfEquipment.Where(i => equipmentName.IndexOf(i) >= 0).FirstOrDefault();
                if (reasonModel != null)
                {
                    line = reasonsModel;
                    break;
                }
            }

            if (line == null) return new Reason();

            var reason = line.Reasons.Where(i => !i.WasUsed).FirstOrDefault();
            if (reason == null)
            {
                line.Reasons.ForEach(i => i.WasUsed = false);
                reason = line.Reasons[0];
            }

            reason.WasUsed = true;
            return reason;
        }
    }
}
