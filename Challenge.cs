using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Challenge
{
    class Challenge : System.Object
    {
        public int id { get; set; }
        public string name { get; set; }
        public string requirementId { get; set; }
        public string requirementMsg { get; set; }
        public string rewardType { get; set; }
        public long rewardAmount { get; set; }
        public string counterName { get; set; }
        public long counterGoal { get; set; }
        public string conditionName { get; set; }
        public long conditionValue { get; set; }
        public string slotsCondition { get; set; }
        public override bool Equals(System.Object obj)
        {
            if (obj == null)
                return false;

            Challenge p = obj as Challenge;
            if ((System.Object)p == null)
                return false;

            return (name == p.name) && (requirementId == p.requirementId) && (requirementMsg == p.requirementMsg)
                && (rewardType == p.rewardType) && (rewardAmount == p.rewardAmount) && (counterName == p.counterName)
                && (counterGoal == p.counterGoal) && (conditionName == p.conditionName) && (conditionValue == p.conditionValue)
                && (slotsCondition == p.slotsCondition);
        }

        public bool Equals(Challenge p)
        {
            if ((object)p == null)
                return false;

            return (name == p.name) && (requirementId == p.requirementId) && (requirementMsg == p.requirementMsg)
                && (rewardType == p.rewardType) && (rewardAmount == p.rewardAmount) && (counterName == p.counterName)
                && (counterGoal == p.counterGoal) && (conditionName == p.conditionName) && (conditionValue == p.conditionValue)
                && (slotsCondition == p.slotsCondition);
        }
        public override int GetHashCode()
        {
            return name.GetHashCode() ^ requirementId.GetHashCode() ^ requirementMsg.GetHashCode() ^ rewardType.GetHashCode()
                 ^ rewardAmount.GetHashCode() ^ counterName.GetHashCode() ^ counterGoal.GetHashCode()
                 ^ conditionName.GetHashCode() ^ conditionValue.GetHashCode() ^ slotsCondition.GetHashCode();
        }
        public bool ShouldSerializeconditionName()
        {
            return (conditionName != null);
        }
        public bool ShouldSerializeconditionValue()
        {
            return (conditionValue != 0);
        }
        public bool ShouldSerializecounterGoal()
        {
            return (counterGoal != 0);
        }
    }
}
