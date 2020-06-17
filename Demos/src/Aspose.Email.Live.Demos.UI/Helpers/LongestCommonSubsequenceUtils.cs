using System;
using System.Collections.Generic;

namespace Aspose.Email.Live.Demos.UI.Helpers
{
    /// <summary>
    /// Longest Common Subsequence algorithm implementation
    /// </summary>
    public static class LongestCommonSubsequenceUtils
    {
        /// <summary>
        /// Build Matrix for Longest Common Subsequence (LCS)
        /// </summary>
        public static int[,] GetMatrix<T>(T[] sequence1, T[] sequence2, IEqualityComparer<T> comparer = null)
        {
            if (comparer == null)
                comparer = EqualityComparer<T>.Default;

            int[,] lcsMatrix = new int[sequence1.Length + 1, sequence2.Length + 1];
            T letter1, letter2;

            for (int i = 0; i < sequence1.Length; i++)
            {
                for (int j = 0; j < sequence2.Length; j++)
                {
                    letter1 = sequence1[i];
                    letter2 = sequence2[j];

                    if (comparer.Equals(letter1, letter2))
                    {
                        if ((i == 0) || (j == 0))
                            lcsMatrix[i, j] = 1;
                        else
                            lcsMatrix[i, j] = 1 + lcsMatrix[i - 1, j - 1];
                    }
                    else
                    {
                        if ((i == 0) && (j == 0))
                            lcsMatrix[i, j] = 0;
                        else if ((i == 0) && !(j == 0))
                            lcsMatrix[i, j] = Math.Max(0, lcsMatrix[i, j - 1]);
                        else if (!(i == 0) && (j == 0))
                            lcsMatrix[i, j] = Math.Max(lcsMatrix[i - 1, j], 0);
                        else if (!(i == 0) && !(j == 0))
                            lcsMatrix[i, j] = Math.Max(lcsMatrix[i - 1, j], lcsMatrix[i, j - 1]);
                    }
                }
            }

            return lcsMatrix;
        }

		/// <summary>
        /// Inflate list with changes for get second sequence from first
        /// </summary>
        public static void ListDiffTreeFromBacktrackMatrix<T>(IList<LSTnode<T>> listToInflate, int[,] lcsMatrix, T[] sequence1, T[] sequence2, IEqualityComparer<T> comparer = null)
        {
            if (comparer == null)
                comparer = EqualityComparer<T>.Default;

            var stack = new Stack<LSTnode<T>>();

            InflateStack(stack, lcsMatrix, sequence1, sequence2, sequence1.Length, sequence2.Length, comparer);

            while (stack.Count > 0)
                listToInflate.Add(stack.Pop());
        }

        static void InflateStack<T>(Stack<LSTnode<T>> stack, int[,] lcsMatrix, T[] sequence1, T[] sequence2, int i, int j, IEqualityComparer<T> comparer)
        {
			// Variable for switch LCS matrix movement direction (to left by columns or to top by rows).
			// Has true by default to show removed node first (better for human reading) in resulted sequence.
            bool shouldGoLeft = true;

            // for additional cycling
            int offset;

            // to traverce rows and columns
            int rowIndex;
            int columnIndex;

			// use goto operator to avoid recursion with stackoverflow

        START:
            if (i > 0 && j > 0 && comparer.Equals(sequence1[i - 1], sequence2[j - 1]))
            {
                stack.Push(new LSTnode<T>() { Mode = NodeMode.NotChanged, Value = sequence1[i - 1] });
                i--;
                j--;
                goto START;
            }
            else
            {
                var shouldRemove = i > 0 && (j == 0 || lcsMatrix[i, j - 1] <= lcsMatrix[i - 1, j]);
                var shouldAdd = j > 0 && (i == 0 || lcsMatrix[i, j - 1] >= lcsMatrix[i - 1, j]);

				// we should select the best direction now by finding the longest added or removed sequence in current position
                if (shouldRemove && shouldAdd)
                {
                    offset = 1;
                    while (true)
                    {
                        columnIndex = j - 1 - offset;
                        rowIndex = i - 1 - offset;

                        if (columnIndex < 0)
                        {
                            shouldGoLeft = false;
                            break;
                        }

                        if (rowIndex < 0)
                        {
                            shouldGoLeft = true;
                            break;
                        }

                        if (lcsMatrix[i, columnIndex] != lcsMatrix[rowIndex, j])
                        {
                            shouldGoLeft = lcsMatrix[i, j - 1 - offset] < lcsMatrix[i - 1 - offset, j];
                            break;
                        }

                        offset++;
                    }
                }

                if (shouldGoLeft)
                {
                    if (shouldRemove)
                    {
                        stack.Push(new LSTnode<T>() { Mode = NodeMode.Removed, Value = sequence1[i - 1] });
                        i--;
                        goto START;
                    }

                    if (shouldAdd)
                    {
                        stack.Push(new LSTnode<T>() { Mode = NodeMode.Added, Value = sequence2[j - 1] });
                        j--;
                        goto START;
                    }
                }
                else
                {
                    if (shouldAdd)
                    {
                        stack.Push(new LSTnode<T>() { Mode = NodeMode.Added, Value = sequence2[j - 1] });
                        j--;
                        goto START;
                    }

                    if (shouldRemove)
                    {
                        stack.Push(new LSTnode<T>() { Mode = NodeMode.Removed, Value = sequence1[i - 1] });
                        i--;
                        goto START;
                    }
                }
            }
        }


        public struct LSTnode<T>
        {
            public NodeMode Mode;
            public T Value;
        }

        public enum NodeMode
        {
            NotChanged = 0,
            Added = 1,
            Removed = 2
        }
    }
}
