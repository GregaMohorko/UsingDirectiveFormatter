﻿namespace Microsoft.VisualStudio.Text
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using UsingDirectiveFormatter.Commands;
    using UsingDirectiveFormatter.Contracts;
    using UsingDirectiveFormatter.Utilities;

    /// <summary>
    /// TextBufferExtensions
    /// </summary>
    public static class TextBufferExtensions
    {
        /// <summary>
        /// The using namespace directive prefix
        /// </summary>
        private static readonly string UsingNamespaceDirectivePrefix = "using";

        /// <summary>
        /// The namespace declaration prefix
        /// </summary>
        private static readonly string NamespaceDeclarationPrefix = "namespace";

        /// <summary>
        /// Formats the specified buffer.
        /// </summary>
        /// <param name="buffer">The buffer.</param>
        public static void Format(this ITextBuffer buffer, FormatOptionGrid options)
        {
            ArgumentGuard.ArgumentNotNull(buffer, "buffer");
            ArgumentGuard.ArgumentNotNull(options, "options");

            // Parse options
            var sortStandards = new List<SortStandard> { options.SortOrderOption, options.ChainedSortOrderOption };
            var insideNamespace = options.InsideNamespace;
            var sortGroups = options.SortGroups.ToList();

            var snapShot = buffer.CurrentSnapshot;

            // Stop processing if there's nothing to process
            if (snapShot.Length == 0)
            {
                return;
            }

            int cursor = 0;
            int tail = 0;

            // using directive related
            var usingDirectives = new List<string>();
            int nsInnerStartPos = 0;
            int nsOuterStartPos = 0;

            // Namespace related flags
            bool nsReached = false;
            int startPos = -1;
            int prensSpanStartPos = 0;

            string indent = "";

            // Using directives before namespace (if any)
            Span? prensSpan = null;
            // Using directives inside namespace, or all usings if there's no namespace
            Span? nsSpan = null;

            bool lastSpanContainsComment = false;
            int spanToPreserve = 0;

            foreach (var line in snapShot.Lines)
            {
                var lineText = line.GetText();
                var lineTextTrimmed = lineText.TrimStart();

                cursor = tail;
                tail += line.LengthIncludingLineBreak;

                if (nsReached && insideNamespace &&
                    !string.IsNullOrWhiteSpace(lineTextTrimmed) &&
                    string.IsNullOrEmpty(indent))
                {
                    indent = lineText.Substring(0, lineText.IndexOf(lineTextTrimmed));
                }

                if (lineTextTrimmed.StartsWith("/", StringComparison.Ordinal))
                {
                    spanToPreserve += line.LengthIncludingLineBreak;
                    lastSpanContainsComment = true;
                }
                else if (string.IsNullOrWhiteSpace(lineTextTrimmed))
                {
                    spanToPreserve += line.LengthIncludingLineBreak;
                }
                else
                {
                    if (lastSpanContainsComment)
                    {
                        // Reset start pos if there are header comments
                        if (prensSpanStartPos == 0 && !nsReached && !usingDirectives.Any())
                        {
                            prensSpanStartPos = spanToPreserve;
                            spanToPreserve = 0;
                        }
                    }
                    else
                    {
                        spanToPreserve = 0;
                    }

                    lastSpanContainsComment = false;

                    if (lineTextTrimmed.StartsWith(UsingNamespaceDirectivePrefix, StringComparison.Ordinal))
                    {
                        if (nsInnerStartPos == 0)
                        {
                            nsInnerStartPos = cursor;
                        }

                        if (startPos < 0)
                        {
                            startPos = cursor;
                        }

                        if (nsOuterStartPos == 0 && !nsReached)
                        {
                            nsOuterStartPos = cursor;
                        }

                        usingDirectives.Add(lineTextTrimmed);
                    }
                    else if (lineTextTrimmed.StartsWith(NamespaceDeclarationPrefix, StringComparison.Ordinal))
                    {
                        if (!nsReached)
                        {
                            prensSpan = new Span(prensSpanStartPos, cursor - prensSpanStartPos - spanToPreserve);
                            nsReached = true;
                        }
                        nsInnerStartPos = tail;
                        startPos = tail;
                        nsOuterStartPos = cursor;
                    }
                    else if (lineTextTrimmed.Equals("{", StringComparison.Ordinal))
                    {
                        if (nsReached)
                        {
                            nsInnerStartPos = tail;
                            startPos = tail;
                        }
                    }
                    else if (lineTextTrimmed.Equals(";", StringComparison.Ordinal))
                    {
                        continue;
                    }
                    else
                    {
                        startPos = Math.Max(startPos, 0);
                        // Special case for when there is no namespace decalartion
                        spanToPreserve = nsReached ? spanToPreserve : spanToPreserve + prensSpanStartPos;

                        nsSpan =
                            new Span(startPos, cursor - startPos - spanToPreserve);
                        break;
                    }
                }
            }

            if (usingDirectives.Any())
            {
                usingDirectives = usingDirectives.Select(s => s.TrimEnd()).OrderBySortStandards(sortStandards).Select(s => indent + s).ToList();
                usingDirectives = usingDirectives.GroupBySortGroups(sortGroups, options.NewLineBetweenSortGroups).ToList();

                var insertPos = nsReached && insideNamespace ? nsInnerStartPos : nsOuterStartPos;
                var insertString = string.Join(Environment.NewLine, usingDirectives) + Environment.NewLine;

                // Testing
                var edit = buffer.CreateEdit();
                edit.Insert(insertPos, insertString);
                if (nsSpan != null)
                {
                    edit.Delete(nsSpan.Value);
                }
                if (prensSpan != null)
                {
                    edit.Delete(prensSpan.Value);
                }
                edit.Apply();
            }
        }
    }
}