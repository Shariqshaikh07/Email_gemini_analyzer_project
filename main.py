import os
import win32com.client
import google.generativeai as genai
# Configure the Gemini API using an environment variable
genai.configure(api_key=os.getenv("GEMINI_API_KEY"))
def read_emails_from_pst(pst_path, max_emails=50):
 """
 Reads a PST file and extracts basic information from a limited number of emails.
 """
 outlook = win32com.client.Dispatch("Outlook.Application")
 namespace = outlook.GetNamespace("MAPI")
 namespace.AddStore(pst_path)
 inbox = namespace.GetDefaultFolder(6) # Inbox folder
 messages = inbox.Items
 messages.Sort("[ReceivedTime]", True) # Newest first
 emails = []
 count = 0
 for message in messages:
 if count >= max_emails:
 break
 try:
 emails.append({
 'subject': message.Subject,
 'body': message.Body[:500], # Limit body length
 'sender': message.SenderName,
 'date': message.ReceivedTime
 })
 count += 1
 except Exception as e:
 print(f"Error reading email: {e}")
 return emails
def build_prompt_from_emails(emails, question):
 """
 Creates a prompt for the Gemini model using extracted email data.
 """
 email_summaries = []
 for i, email in enumerate(emails, 1):
 summary = (
 f"Email {i}\n"
 f"Subject: {email['subject']}\n"
 f"Sender: {email['sender']}\n"
 f"Date: {email['date']}\n"
 f"Body: {email['body']}\n"
 )
 email_summaries.append(summary)
 combined = "\n---\n".join(email_summaries)
 prompt = (
 f"You have access to the following email records:\n\n{combined}\n\n"
 f"Based on this information, answer the following question:\n{question}"
 )
 return prompt
def ask_gemini(question, emails):
 """
 Sends the prompt to Gemini and returns the response.
 """
 prompt = build_prompt_from_emails(emails, question)
 try:
 model = genai.GenerativeModel("gemini-1.5-flash")
 response = model.generate_content([{"text": prompt}])
 return response.text.strip()
 except Exception as e:
 return f"Error querying Gemini: {e}"
def main():
 """
 Main program flow: load emails and interactively ask questions.
 """
 pst_path = input("Enter the full path to your PST file: ").strip()
 if not os.path.exists(pst_path):
 print("Invalid PST file path.")
 return
 print("Loading emails...")
 emails = read_emails_from_pst(pst_path)
 print(f"{len(emails)} emails loaded.")
 while True:
 question = input("\nAsk a question about the emails (type 'exit' to quit):
").strip()
 if question.lower() == 'exit':
 print("Goodbye!")
 break
 answer = ask_gemini(question, emails)
 print("\nGemini's Answer:")
 print(answer)
if __name__ == "__main__":
 main()
