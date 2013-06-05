using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Fusion.Messages.SocialCare.Schemas.Patterns
{
    public interface IDatedObject
    {
        string auditUserName
        {
            get;
            set;
        }

        System.DateTime effectiveFrom
        {
            get;
            set;
        }

        bool effectiveFromSpecified
        {
            get;
            set;
        }

        System.DateTime effectiveTo
        {
            get;
            set;
        }

        bool effectiveToSpecified
        {
            get;
            set;
        }
    }

    public static class DatedObjectExtensions
    {
        public static DateTime? GetEffectiveFrom(this IDatedObject datedObject)
        {
            if (datedObject == null)
            {
                return null;
            }

            return datedObject.effectiveFromSpecified
                ? datedObject.effectiveFrom
                : (DateTime?)null;
        }

        public static DateTime? GetEffectiveTo(this IDatedObject datedObject)
        {
            if (datedObject == null)
            {
                return null;
            }

            return datedObject.effectiveToSpecified
                ? datedObject.effectiveTo
                : (DateTime?)null;
        }

        /// <summary>
        /// Does not include the future objects. Chooses the latest one which started to be effective now or in the past and is still effective now.
        /// </summary>
        /// <typeparam name="TDatedObject"></typeparam>
        /// <param name="datedObjects"></param>
        /// <returns></returns>
        public static TDatedObject GetCurrentlyEffective<TDatedObject>(this IEnumerable<TDatedObject> datedObjects)
            where TDatedObject : class, IDatedObject
        {
            if (datedObjects == null)
            {
                return null;
            }

            var now = DateTime.Now;//ConnectorContext.RealTimeClock.Now; 

            var antiChronologically = from datedObject in datedObjects
                                      where !datedObject.effectiveFromSpecified //Senders that do not support dated objects: this means "as of now"
                                        || (
                                            datedObject.effectiveFromSpecified && datedObject.effectiveFrom <= now //Specified and not in future
                                            && (!datedObject.effectiveToSpecified || datedObject.effectiveTo > now) //Open-ended or ends now or in future.
                                           )
                                      orderby datedObject.effectiveFrom descending //latest effective first, if not specified, then lands at the end of the list.
                                      select datedObject;

            return antiChronologically.FirstOrDefault();
        }

        /// <summary>
        /// Chooses the latest one which started to be effective now or in the past and is still effective now.
        /// If no such one is found, then returns the one which will become effective in the nearest future.
        /// </summary>
        /// <typeparam name="TDatedObject"></typeparam>
        /// <param name="datedObjects"></param>
        /// <returns></returns>
        public static TDatedObject GetCurrentlyEffectiveOrNearestInFuture<TDatedObject>(this IEnumerable<TDatedObject> datedObjects)
            where TDatedObject : class, IDatedObject
        {
            if (datedObjects == null)
            {
                return null;
            }

            var now = DateTime.Now; //ConnectorContext.RealTimeClock.Now;

            var notFinished = from datedObject in datedObjects
                              where !datedObject.effectiveToSpecified
                                || (datedObject.effectiveToSpecified && datedObject.effectiveTo > now)
                              select datedObject;
            notFinished = notFinished.ToArray();

            var lastCurrent = (from datedObject in notFinished
                              where datedObject.effectiveFromSpecified && datedObject.effectiveFrom <= now
                              orderby datedObject.effectiveFrom descending
                              select datedObject).FirstOrDefault();

            if (lastCurrent != null)
            {
                //We've got one which is current:
                return lastCurrent;
            }

            //Otherwise we need to find the first which is future.
            //The ones that have no start date are deemed to be the most away in time:

            var futureOnes = notFinished; //Note that we already know from the previous query and if, that this contains only future objects.

            var firstFuture = (from datedObject in futureOnes
                              orderby datedObject.effectiveFromSpecified ? datedObject.effectiveFrom : DateTime.MaxValue ascending
                              select datedObject).FirstOrDefault();

            //This can be null if no future objects are on the list.
            return firstFuture;
        }
    }
}
