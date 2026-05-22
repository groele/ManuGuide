using System;
using System.ComponentModel;

namespace Manuscript_guide.Models
{
    /// <summary>
    /// Represents a single diagnostic issue found in the Word document.
    /// Implements INotifyPropertyChanged so that UI card states (Normal vs Corrected)
    /// can be toggled via data binding inside ItemsControl DataTemplates.
    /// </summary>
    public class IssueItem : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private bool _isFixed = false;
        private bool _isIgnored = false;

        public string IssueId { get; set; }
        public string Type { get; set; }        // e.g., "punc", "cap", "dash", "data", "ital", "sub", "word"
        public string Subtype { get; set; }     // e.g., "UndefinedAcronym", "CasingInconsistency"
        public string RuleId { get; set; }      // Stable rule id, e.g., "sub.element_formula_subscript"
        public string RuleTitle { get; set; }   // User-facing rule name
        public string MatchReason { get; set; } // Short explanation of why the rule matched
        public int Start { get; set; }          // Character index start in Word range
        public int End { get; set; }            // Character index end in Word range
        public string OriginalText { get; set; }
        public string RecommendFix { get; set; }
        public string Desc { get; set; }        // Diagnostic details
        public string Context { get; set; }     // Context snippet around the issue
        public string TargetMark { get; set; }  // Highlight element identifier

        /// <summary>
        /// Whether this issue has been accepted and corrected.
        /// Drives the Visibility of the success card via binding.
        /// </summary>
        public bool IsFixed
        {
            get => _isFixed;
            set
            {
                if (_isFixed != value)
                {
                    _isFixed = value;
                    OnPropertyChanged(nameof(IsFixed));
                    OnPropertyChanged(nameof(NormalStateVisibility));
                    OnPropertyChanged(nameof(CorrectedStateVisibility));
                    OnPropertyChanged(nameof(CardVisibility));
                }
            }
        }

        /// <summary>
        /// Whether this issue has been dismissed/ignored.
        /// </summary>
        public bool IsIgnored
        {
            get => _isIgnored;
            set
            {
                if (_isIgnored != value)
                {
                    _isIgnored = value;
                    OnPropertyChanged(nameof(IsIgnored));
                    OnPropertyChanged(nameof(CardVisibility));
                }
            }
        }

        // Computed visibility properties for XAML binding
        public System.Windows.Visibility NormalStateVisibility
            => (_isFixed) ? System.Windows.Visibility.Collapsed : System.Windows.Visibility.Visible;

        public System.Windows.Visibility CorrectedStateVisibility
            => (_isFixed) ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;

        public System.Windows.Visibility CardVisibility
            => (_isIgnored) ? System.Windows.Visibility.Collapsed : System.Windows.Visibility.Visible;

        public IssueItem()
        {
            IssueId = Guid.NewGuid().ToString();
        }

        protected void OnPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
