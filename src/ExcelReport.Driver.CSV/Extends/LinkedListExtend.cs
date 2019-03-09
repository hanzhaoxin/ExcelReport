using System.Collections.Generic;

namespace ExcelReport.Driver.CSV.Extends
{
    internal static class LinkedListExtend
    {
        public static LinkedList<T> AppendAll<T>(this LinkedList<T> self, LinkedList<T> list)
        {
            foreach (var item in list)
            {
                self.AddLast(item);
            }
            return self;
        }

        public static LinkedList<T> RemoveAll<T>(this LinkedList<T> self, LinkedList<T> list)
        {
            foreach (var item in list)
            {
                self.Remove(item);
            }
            return self;
        }
    }
}