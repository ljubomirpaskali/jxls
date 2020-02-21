package jxls.example;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.List;

import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.core.io.ClassPathResource;

@SpringBootApplication
public class JxlsExample {

	private static final String LOGO = "logo.png";
	private static final String TEMPLATE = "template.xlsx";
	private static final String TIMESHEET = "timesheet.xlsx";

	public static void main(String[] args) throws IOException {
		JxlsHelper.getInstance().processTemplate(getTemplateStream(), getTimeSheetStream(), getContext());
	}

	private static InputStream getTemplateStream() throws IOException {
		return new ClassPathResource(TEMPLATE).getInputStream();
	}

	private static OutputStream getTimeSheetStream() throws FileNotFoundException {
		return new FileOutputStream(TIMESHEET);
	}

	private static Context getContext() throws IOException {
		Context context = new Context();
		context.putVar("logo", getLogoStream().readAllBytes());
		context.putVar("employee", getEmployee());
		context.putVar("client", getClient());
		context.putVar("worklogEntries", getWorklogEntries());
		return context;
	}

	private static InputStream getLogoStream() throws IOException {
		return new ClassPathResource(LOGO).getInputStream();
	}

	private static User getEmployee() {
		return new User("Vladan Bjelic", "vladan.bjelic@hybrid-it.rs");
	}

	private static Client getClient() {
		return new Client("Luxs Insights", "Kris Spitsbaard");
	}

	private static List<WorklogEntry> getWorklogEntries() {
		return Arrays.asList(
			new WorklogEntry("Monday", "17.02.2020.", "Luxs FE", "development", 1, 7),
			new WorklogEntry("Tuesday", "18.02.2020.", "Luxs BE", "testing", 2, 6),
			new WorklogEntry("Wednesday", "19.02.2020.", "Luxs FE", "development", 3, 5),
			new WorklogEntry("Thursday", "20.02.2020.", "Luxs BE", "testing", 4, 4),
			new WorklogEntry("Friday", "21.02.2020.", "Luxs FE", "development", 5, 3)
		);
	}

	public static class User {
		private final String name;
		private final String email;

		public User(String name, String email) {
			this.name = name;
			this.email = email;
		}

		public String getName() {
			return name;
		}
		public String getEmail() {
			return email;
		}
	}

	public static class Client {
		private final String name;
		private final String contact;

		public Client(String name, String contact) {
			this.name = name;
			this.contact = contact;
		}

		public String getName() {
			return name;
		}
		public String getContact() {
			return contact;
		}
	}

	public static class WorklogEntry {
		private final String day;
		private final String date;
		private final String project;
		private final String activity;
		private final int hours;
		private final int vacation;

		public WorklogEntry(String day, String date, String project, String activity, int hours, int vacation) {
			this.day = day;
			this.date = date;
			this.project = project;
			this.activity = activity;
			this.hours = hours;
			this.vacation = vacation;
		}

		public String getDay() {
			return day;
		}
		public String getDate() {
			return date;
		}
		public String getProject() {
			return project;
		}
		public String getActivity() {
			return activity;
		}
		public int getHours() {
			return hours;
		}
		public int getVacation() {
			return vacation;
		}
	}

}
