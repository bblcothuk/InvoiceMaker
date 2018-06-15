import java.awt.*;
import javax.swing.*;
import javax.sound.midi.*;
import java.util.*;
import java.awt.event.*;

public class InvoiceMakerGUI {

	JPanel mainPanel;
	JFrame theFrame;

	public void buildGUI() {
		theFrame = new JFrame("Invoice Maker");
		theFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		BorderLayout layout = new BorderLayout();
		JPanel background = new JPanel(layout);
		background.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

		Box buttonBox = new Box(BoxLayout.Y_AXIS);

		JButton start = new JButton("Start");
		start.addActionListener(new MyStartListener());
		buttonBox.add(start);

		JButton stop = new JButton("Stop");
		stop.addActionListener(new MyStopListener());
		buttonBox.add(stop);

		Box nameBox = new Box(BoxLayout.Y_AXIS);

		background.add(BorderLayout.CENTER, buttonBox);

		theFrame.getContentPane().add(background);

		theFrame.setBounds(50, 50, 300, 300);
		theFrame.pack();
		theFrame.setVisible(true);
	}

	public class MyStartListener implements ActionListener {

		public void actionPerformed(ActionEvent a) {
		}
	}

	public class MyStopListener implements ActionListener {

		public void actionPerformed(ActionEvent a) {
		}
	}

}
