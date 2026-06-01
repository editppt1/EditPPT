
%======================================================================
\section{Limitations of Win32COM}\label{sec:win32com}
%======================================================================

\begin{itemize}
  \item \textbf{Difficulty with parallel processing:} Win32COM does not natively support parallel execution, causing speed bottlenecks when processing a large number of slides.
  \item \textbf{TextFrame / TextFrame2 compatibility issues}
  \item \textbf{Equation (OMath) handling issues}

  \begin{figure}[H]
    \centering
    \includegraphics[width=0.7\textwidth]{equation handling-1.png}
    \caption{Equation Handling Issue}
  \end{figure}

  \begin{figure}[H]
    \centering
    \includegraphics[width=0.7\textwidth]{equation form-1.png}
    \caption{Equation Form}
  \end{figure}

  \item \textbf{Paragraph() handling issues}

  \begin{figure}[H]
    \centering
    \includegraphics[width=0.7\textwidth]{paragraph_with_carrigereturn-1.png}
    \caption{Paragraph with Carriage Return}
  \end{figure}

  \begin{figure}[H]
    \centering
    \includegraphics[width=0.7\textwidth]{broken_paragraph -1.png}
    \caption{Broken Paragraph}
  \end{figure}

  Resolution:
  \begin{minted}{python}
# Previous approach
paragraph.Text = ""
paragraph.InsertAfter("new text")
# Current approach
full_text = "".join(mapped_run_texts)
paragraph.Text = full_text
  \end{minted}
\end{itemize}


%======================================================================
\section{Limitations of Claude PPT's Editing Approach}\label{sec:claude}
%======================================================================

\begin{itemize}
  \item The current Claude PPT reads and modifies the OOXML directly for all slides, resulting in excessive token costs and processing time.
  \item This is particularly inefficient for simple repetitive editing tasks.
  \item Furthermore, reading the full OOXML verbatim causes excessive hallucinations when modifying text with finely separated Run-level formatting, leading to observed issues such as broken spacing.
\end{itemize}