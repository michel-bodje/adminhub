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

// lawyers is a Promise that resolves to an array of Lawyer objects
export const lawyers = loadLawyers();

/**
 * Returns an array of all lawyers.
 * @returns {Array<Lawyer>} - An array of all lawyer objects.
 */
export async function getAllLawyers() {
  return await lawyers;
}

/**
 * Returns the lawyer object with the given ID.
 * @param {string} lawyerId - The unique ID of the lawyer.
 * @returns {Lawyer | null} - The lawyer object if found, null if not.
 */
export async function getLawyer(lawyerId) {
  try {
    const allLawyers = await lawyers;
    const lawyer = allLawyers.find((l) => l.id === lawyerId);
    if (!lawyer) {
      throw new Error(`Lawyer with ID ${lawyerId} not found.`);
    }
    return lawyer;
  } catch (error) {
    console.warn(error.message);
    return null;
  }
}
