namespace ExceltoXML
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.S2KFunctional = new System.Windows.Forms.CheckBox();
            this.Interlocking = new System.Windows.Forms.CheckBox();
            this.ASCV = new System.Windows.Forms.CheckBox();
            this.Switch = new System.Windows.Forms.CheckBox();
            this.Route = new System.Windows.Forms.CheckBox();
            this.Cycle = new System.Windows.Forms.CheckBox();
            this.Overlap = new System.Windows.Forms.CheckBox();
            this.MBL = new System.Windows.Forms.CheckBox();
            this.PSAPR = new System.Windows.Forms.CheckBox();
            this.Signal = new System.Windows.Forms.CheckBox();
            this.Point = new System.Windows.Forms.CheckBox();
            this.ESP = new System.Windows.Forms.CheckBox();
            this.TrafficDirection = new System.Windows.Forms.CheckBox();
            this.Subroute = new System.Windows.Forms.CheckBox();
            this.TrackCircuit = new System.Windows.Forms.CheckBox();
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Alstom", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(114, 58);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(150, 42);
            this.button1.TabIndex = 0;
            this.button1.Text = "Generate Output";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("Alstom", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.Location = new System.Drawing.Point(539, 11);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(122, 40);
            this.button2.TabIndex = 0;
            this.button2.Text = "Browse Input";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(114, 12);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(418, 40);
            this.textBox1.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(57, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "File Name:";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("Alstom", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button3.Location = new System.Drawing.Point(270, 58);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(150, 42);
            this.button3.TabIndex = 3;
            this.button3.Text = "Open Folder";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.S2KFunctional);
            this.panel1.Controls.Add(this.Interlocking);
            this.panel1.Controls.Add(this.ASCV);
            this.panel1.Controls.Add(this.Switch);
            this.panel1.Controls.Add(this.Route);
            this.panel1.Controls.Add(this.Cycle);
            this.panel1.Controls.Add(this.Overlap);
            this.panel1.Controls.Add(this.MBL);
            this.panel1.Controls.Add(this.PSAPR);
            this.panel1.Controls.Add(this.Signal);
            this.panel1.Controls.Add(this.Point);
            this.panel1.Controls.Add(this.ESP);
            this.panel1.Controls.Add(this.TrafficDirection);
            this.panel1.Controls.Add(this.Subroute);
            this.panel1.Controls.Add(this.TrackCircuit);
            this.panel1.Location = new System.Drawing.Point(114, 106);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(546, 230);
            this.panel1.TabIndex = 4;
            // 
            // S2KFunctional
            // 
            this.S2KFunctional.AutoSize = true;
            this.S2KFunctional.Location = new System.Drawing.Point(273, 188);
            this.S2KFunctional.Name = "S2KFunctional";
            this.S2KFunctional.Size = new System.Drawing.Size(95, 17);
            this.S2KFunctional.TabIndex = 14;
            this.S2KFunctional.Text = "S2KFunctional";
            this.S2KFunctional.UseVisualStyleBackColor = true;
            this.S2KFunctional.CheckedChanged += new System.EventHandler(this.S2KFunctional_CheckedChanged);
            // 
            // Interlocking
            // 
            this.Interlocking.AutoSize = true;
            this.Interlocking.Location = new System.Drawing.Point(273, 147);
            this.Interlocking.Name = "Interlocking";
            this.Interlocking.Size = new System.Drawing.Size(81, 17);
            this.Interlocking.TabIndex = 13;
            this.Interlocking.Text = "Interlocking";
            this.Interlocking.UseVisualStyleBackColor = true;
            this.Interlocking.CheckedChanged += new System.EventHandler(this.Interlocking_CheckedChanged);
            // 
            // ASCV
            // 
            this.ASCV.AutoSize = true;
            this.ASCV.Location = new System.Drawing.Point(273, 104);
            this.ASCV.Name = "ASCV";
            this.ASCV.Size = new System.Drawing.Size(54, 17);
            this.ASCV.TabIndex = 12;
            this.ASCV.Text = "ASCV";
            this.ASCV.UseVisualStyleBackColor = true;
            this.ASCV.CheckedChanged += new System.EventHandler(this.ASCV_CheckedChanged);
            // 
            // Switch
            // 
            this.Switch.AutoSize = true;
            this.Switch.Location = new System.Drawing.Point(273, 62);
            this.Switch.Name = "Switch";
            this.Switch.Size = new System.Drawing.Size(58, 17);
            this.Switch.TabIndex = 11;
            this.Switch.Text = "Switch";
            this.Switch.UseVisualStyleBackColor = true;
            this.Switch.CheckedChanged += new System.EventHandler(this.Switch_CheckedChanged);
            // 
            // Route
            // 
            this.Route.AutoSize = true;
            this.Route.Location = new System.Drawing.Point(273, 19);
            this.Route.Name = "Route";
            this.Route.Size = new System.Drawing.Size(55, 17);
            this.Route.TabIndex = 10;
            this.Route.Text = "Route";
            this.Route.UseVisualStyleBackColor = true;
            this.Route.CheckedChanged += new System.EventHandler(this.Route_CheckedChanged);
            // 
            // Cycle
            // 
            this.Cycle.AutoSize = true;
            this.Cycle.Location = new System.Drawing.Point(166, 147);
            this.Cycle.Name = "Cycle";
            this.Cycle.Size = new System.Drawing.Size(52, 17);
            this.Cycle.TabIndex = 9;
            this.Cycle.Text = "Cycle";
            this.Cycle.UseVisualStyleBackColor = true;
            this.Cycle.CheckedChanged += new System.EventHandler(this.Cycle_CheckedChanged);
            // 
            // Overlap
            // 
            this.Overlap.AutoSize = true;
            this.Overlap.Location = new System.Drawing.Point(166, 62);
            this.Overlap.Name = "Overlap";
            this.Overlap.Size = new System.Drawing.Size(63, 17);
            this.Overlap.TabIndex = 8;
            this.Overlap.Text = "Overlap";
            this.Overlap.UseVisualStyleBackColor = true;
            this.Overlap.CheckedChanged += new System.EventHandler(this.Overlap_CheckedChanged);
            // 
            // MBL
            // 
            this.MBL.AutoSize = true;
            this.MBL.Location = new System.Drawing.Point(166, 188);
            this.MBL.Name = "MBL";
            this.MBL.Size = new System.Drawing.Size(48, 17);
            this.MBL.TabIndex = 7;
            this.MBL.Text = "MBL";
            this.MBL.UseVisualStyleBackColor = true;
            this.MBL.CheckedChanged += new System.EventHandler(this.MBL_CheckedChanged);
            // 
            // PSAPR
            // 
            this.PSAPR.AutoSize = true;
            this.PSAPR.Location = new System.Drawing.Point(166, 104);
            this.PSAPR.Name = "PSAPR";
            this.PSAPR.Size = new System.Drawing.Size(62, 17);
            this.PSAPR.TabIndex = 6;
            this.PSAPR.Text = "PSAPR";
            this.PSAPR.UseVisualStyleBackColor = true;
            this.PSAPR.CheckedChanged += new System.EventHandler(this.PSAPR_CheckedChanged);
            // 
            // Signal
            // 
            this.Signal.AutoSize = true;
            this.Signal.Location = new System.Drawing.Point(166, 19);
            this.Signal.Name = "Signal";
            this.Signal.Size = new System.Drawing.Size(55, 17);
            this.Signal.TabIndex = 5;
            this.Signal.Text = "Signal";
            this.Signal.UseVisualStyleBackColor = true;
            this.Signal.CheckedChanged += new System.EventHandler(this.Signal_CheckedChanged);
            // 
            // Point
            // 
            this.Point.AutoSize = true;
            this.Point.Location = new System.Drawing.Point(14, 188);
            this.Point.Name = "Point";
            this.Point.Size = new System.Drawing.Size(50, 17);
            this.Point.TabIndex = 4;
            this.Point.Text = "Point";
            this.Point.UseVisualStyleBackColor = true;
            this.Point.CheckedChanged += new System.EventHandler(this.Point_CheckedChanged);
            // 
            // ESP
            // 
            this.ESP.AutoSize = true;
            this.ESP.Location = new System.Drawing.Point(14, 147);
            this.ESP.Name = "ESP";
            this.ESP.Size = new System.Drawing.Size(47, 17);
            this.ESP.TabIndex = 3;
            this.ESP.Text = "ESP";
            this.ESP.UseVisualStyleBackColor = true;
            this.ESP.CheckedChanged += new System.EventHandler(this.ESP_CheckedChanged);
            // 
            // TrafficDirection
            // 
            this.TrafficDirection.AutoSize = true;
            this.TrafficDirection.Location = new System.Drawing.Point(14, 104);
            this.TrafficDirection.Name = "TrafficDirection";
            this.TrafficDirection.Size = new System.Drawing.Size(98, 17);
            this.TrafficDirection.TabIndex = 2;
            this.TrafficDirection.Text = "TrafficDirection";
            this.TrafficDirection.UseVisualStyleBackColor = true;
            this.TrafficDirection.CheckedChanged += new System.EventHandler(this.TrafficDirection_CheckedChanged);
            // 
            // Subroute
            // 
            this.Subroute.AutoSize = true;
            this.Subroute.Location = new System.Drawing.Point(14, 62);
            this.Subroute.Name = "Subroute";
            this.Subroute.Size = new System.Drawing.Size(69, 17);
            this.Subroute.TabIndex = 1;
            this.Subroute.Text = "Subroute";
            this.Subroute.UseVisualStyleBackColor = true;
            this.Subroute.CheckedChanged += new System.EventHandler(this.Subroute_CheckedChanged);
            // 
            // TrackCircuit
            // 
            this.TrackCircuit.AutoSize = true;
            this.TrackCircuit.Location = new System.Drawing.Point(14, 19);
            this.TrackCircuit.Name = "TrackCircuit";
            this.TrackCircuit.Size = new System.Drawing.Size(83, 17);
            this.TrackCircuit.TabIndex = 0;
            this.TrackCircuit.Text = "TrackCircuit";
            this.TrackCircuit.UseVisualStyleBackColor = true;
            this.TrackCircuit.CheckedChanged += new System.EventHandler(this.TrackCircuit_CheckedChanged);
            // 
            // button4
            // 
            this.button4.Font = new System.Drawing.Font("Alstom", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button4.Location = new System.Drawing.Point(114, 342);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(150, 42);
            this.button4.TabIndex = 5;
            this.button4.Text = "Select All";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.button5.Font = new System.Drawing.Font("Alstom", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button5.Location = new System.Drawing.Point(270, 342);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(150, 42);
            this.button5.TabIndex = 6;
            this.button5.Text = "Deselect All";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // Form1
            // 
            this.AutoScroll = true;
            this.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.ClientSize = new System.Drawing.Size(754, 424);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.textBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "Form1";
            this.Text = "U200 IDP GEN";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.CheckBox Point;
        private System.Windows.Forms.CheckBox ESP;
        private System.Windows.Forms.CheckBox TrafficDirection;
        private System.Windows.Forms.CheckBox Subroute;
        private System.Windows.Forms.CheckBox TrackCircuit;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.CheckBox MBL;
        private System.Windows.Forms.CheckBox PSAPR;
        private System.Windows.Forms.CheckBox Signal;
        private System.Windows.Forms.CheckBox ASCV;
        private System.Windows.Forms.CheckBox Switch;
        private System.Windows.Forms.CheckBox Route;
        private System.Windows.Forms.CheckBox Cycle;
        private System.Windows.Forms.CheckBox Overlap;
        private System.Windows.Forms.CheckBox Interlocking;
        private System.Windows.Forms.CheckBox S2KFunctional;
    }
}

