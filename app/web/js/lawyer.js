/**
 * Class representing a lawyer.
 * Lawyers are defined in lawyerData.json
 * @class
 */
export class Lawyer {
  /**
   * Constructor for the Lawyer class.
   * @param {string} id - The unique ID of the lawyer.
   * @param {string} name - The name of the lawyer.
   * @param {string} email - The email address of the lawyer.
   * @param {{start: string, end: string}} workingHours - The object containing the start and end time of the working hours of the lawyer.
   * @param {number} breakMinutes - The number of minutes of break time allocated to the lawyer.
   * @param {number} maxDailyAppointments - The maximum number of appointments the lawyer can have in a day.
   * @param {Array<string>} specialties - An array of the specialties of the lawyer.
   */
  constructor(id, name, email, workingHours, breakMinutes, maxDailyAppointments, specialties) {
    this.id = id;
    this.name = name;
    this.email = email;
    this.workingHours = workingHours;
    this.breakMinutes = breakMinutes;
    this.maxDailyAppointments = maxDailyAppointments;
    this.specialties = specialties;
  }
}

// Create lawyer objects from the data
export async function loadLawyers() {
  const data = await window.pywebview.api.get_lawyers();
  if (data.error) {
    console.error("Failed to load lawyers:", data.error);
    return [];
  }
  return data.map(lawyer => new Lawyer(
    lawyer.id,
    lawyer.name,
    lawyer.email,
    lawyer.workingHours,
    lawyer.breakMinutes,
    lawyer.maxDailyAppointments,
    lawyer.specialties
  ));
}

let lawyerList = [];

export function setLawyers(list) {
  lawyerList = list;
}

export function getLawyers() {
  return lawyerList;
}

export function getLawyerById(id) {
  return lawyerList.find(l => l.id === id) || null;
}
