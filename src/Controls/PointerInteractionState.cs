using System.Windows.Forms;

namespace ExcelTrainingMonitor.Controls
{
    internal sealed class PointerInteractionState
    {
        public bool Hovering { get; private set; }
        public bool Pressing { get; private set; }

        public void Attach(Control control)
        {
            control.MouseEnter += OnMouseEnter;
            control.MouseLeave += OnMouseLeave;
            control.MouseDown += OnMouseDown;
            control.MouseUp += OnMouseUp;
        }

        private void OnMouseEnter(object sender, System.EventArgs e)
        {
            var control = (Control)sender;
            Hovering = true;
            control.Invalidate();
        }

        private void OnMouseLeave(object sender, System.EventArgs e)
        {
            var control = (Control)sender;
            Hovering = false;
            Pressing = false;
            control.Invalidate();
        }

        private void OnMouseDown(object sender, MouseEventArgs e)
        {
            Pressing = true;
            ((Control)sender).Invalidate();
        }

        private void OnMouseUp(object sender, MouseEventArgs e)
        {
            Pressing = false;
            ((Control)sender).Invalidate();
        }
    }
}
