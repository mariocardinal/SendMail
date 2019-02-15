using System;
using System.Text;
using System.Net.Mail;
using System.Net.Mime;
using CsvHelper;
using System.IO;

namespace sendmail
{
    public class Recipient
    {
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string Email { get; set; }
        public string Group { get; set; }
        public string Lang { get; set; }
        public string Contact { get; set; }
        public string Tags { get; set; }
    }
    class Program
    {
        static void Main(string[] args)  {
            string senderName = "Mario Cardinal";
            string senderEmail = "mario.cardinal@to-do.studio";
            string senderPwd = "Enter Password here";
            string partnerName = "Erik Renaud";
            
            using (var reader = new StreamReader("c:\\tmp\\mario_contact.csv", Encoding.UTF8))
            using (var csv = new CsvReader(reader))
            using (
               SmtpClient smtp = new SmtpClient {
                  Host = "smtp.office365.com",
                  Port = 587,
                  Credentials = new System.Net.NetworkCredential(senderEmail, senderPwd),
                  EnableSsl = true
               }
            )
            {
                Random random = new Random();    

                //CSVReader will now read the whole CVS file into an enumerable
                var recipients = csv.GetRecords<Recipient>();
                foreach (Recipient recipient in recipients) 
                {
                    Console.WriteLine("{0}, {1}, {2}, {3}, {4}", recipient.LastName, recipient.FirstName, recipient.Email, recipient.Group, recipient.Lang);
                
                    using (
                        MailMessage message = new MailMessage {
                            To = { new MailAddress(recipient.Email, recipient.FirstName + " " + recipient.LastName) },
                            Sender = new MailAddress(senderEmail, senderName),
                            From = new MailAddress(senderEmail, senderName),
                            IsBodyHtml = true,
                        }
                    )
                    {
                        try { 
                            message.Subject = getSubject(recipient.Lang);
                            message.AlternateViews.Add(getHtmlBody(senderName, partnerName, recipient.FirstName, recipient.Lang));
                            smtp.Send(message);
                        }
                        catch (Exception excp) {
                            Console.Write(excp.Message);
                            Console.ReadKey();
                        }
                    }
                    int msWait = random.Next(60000, 300000);
                    Console.WriteLine("Wait {0} milliseconds", msWait);
                    System.Threading.Thread.Sleep(msWait); // Wait betweeen 1 to 5 min
                } 
            }
        }

        private static string getSubject(string lang) {
            string subject = "Subscribe to my newsletter";
            if(lang == "Fr")
                subject = "Abonne-toi à mon infolettre";
          return subject;
        }
        private static AlternateView getHtmlBody(string senderName, string partnerName, string recipientFirstName, string recipientLang) {
            StringBuilder bodyHtml = (recipientLang == "Fr") ? new StringBuilder(getTemplateFr()) : new StringBuilder(getTemplateEn());
            bodyHtml.Replace("__recipientFirstName__", recipientFirstName);
            bodyHtml.Replace("__senderName__", senderName);
            bodyHtml.Replace("__partnerName__", partnerName);
            LinkedResource marioPhoto = new LinkedResource("c:/tmp/mario_cardinal_100x100_bw.jpg", MediaTypeNames.Image.Jpeg);
            marioPhoto.ContentId = Guid.NewGuid().ToString();
            bodyHtml.Replace("__marioPhotoCid__",  marioPhoto.ContentId);
            AlternateView alternateView = AlternateView.CreateAlternateViewFromString(bodyHtml.ToString(), null, MediaTypeNames.Text.Html);
            alternateView.LinkedResources.Add(marioPhoto);
            return alternateView;
        }

        private static string getTemplateFr() {
            return @"<body> 
            <p>Bonjour __recipientFirstName__,</p>
            <p><img src='cid:__marioPhotoCid__' height='100' width='100' align='left' hspace='10'/>
            Ici __senderName__. Je désire partager mon parcours d'entrepreneur avec toi. C'est pourquoi je t’invite à t’abonner à mon infolettre <a href='https://tinyletter.com/todostudio-fr'>Destination To-Do Studio</a>.
            </p>
            <p>
            L'infolettre est à propos de l'entreprise que j’ai fondé avec __partnerName__. Nous venons de passer plusieurs années à chercher une solution simple et élégante pour répondre au problème suivant: <em>Comment permettre à plusieurs personnes de compléter ensemble une liste de to-do <b>commune</b>?</em> Et ceci sans utiliser l’approche classique qui consiste à coordonner le travail avec un plan et une structure hiérarchique, solution typique qu’on retrouve avec les logiciels de gestion de projets.  Nous voulions une approche plus simple, plus ouverte et reposant sur le principe des équipes autogérées. Ceci afin d’éviter aux leaders la tracasserie des suivis et des pourparlers difficiles qui se produisent lorsqu’on pousse et tire les autres à prendre initiative.
            </p>
            <p>
            Après cinq ans de recherche et développement (R&D), nous sommes maintenant prêts à commercialiser notre innovation: Voici To-Do Studio, un assistant automatisé proposé sous forme d’un service logiciel.
            </p>
            <p>
            Voici trois bénéfices que tu obtiendras en lisant notre infolettre :
                <ol>
                    <li>
                    Tu acquerras un regard privilégié sur le parcours d’une jeune entreprise qui vise à obtenir des milliers d’utilisateurs dans le monde entier;
                    </li>
                    <li>
                    Tu suivras les étapes de commercialisation d’un logiciel et tu pourras ainsi mieux comprendre la technologie;
                    </li>
                    <li>
                    Enfin, tu apprendras à quel point To-Do Studio est unique pour inciter les autres à prendre initiative.
                    </li>
                </ol>
            </p>
            <div style='text-align: center'>
                <div style='display: inline-block'>
                    <!--[if mso]>
                        <v:roundrect xmlns:v='urn:schemas-microsoft-com:vml' xmlns:w='urn:schemas-microsoft-com:office:word' 
                        href='https://tinyletter.com/todostudio-fr' 
                        style='height:60px;v-text-anchor:middle;width:600px;' arcsize='15%' strokecolor='#008000' fillcolor='#008000'>
                            <w:anchorlock/>
                            <center style='color:#ffffff;font-family:sans-serif;font-size:14px;font-weight:bold;'>Oui, je désire suivre votre parcours &rarr;</center>
                        </v:roundrect>
                    <![endif]-->
                    <a href='https://tinyletter.com/todostudio-fr' 
                    style='background-color:#008000;border:1px solid #008000;border-radius:10px;color:#ffffff;display:inline-block;font-family:sans-serif;font-size:14px;font-weight:bold;line-height:60px;text-align:center;text-decoration:none;width:350px;-webkit-text-size-adjust:none;mso-hide:all;'>
                    Oui, je désire suivre votre parcours &rarr;</a>
                </div>
            </div>
            <p>
            L’infolettre est gratuite. Elle est publiée à tous les 2 mois, 6 fois par année. L’abonnement n’exige aucune obligation de ta part. La seule information que tu dois fournir pour t’inscrire est ton adresse courriel (elle ne sera jamais partagée à des tiers). Tu peux te désabonner en tout temps. Si tu désires avoir un aperçu du contenu de l’infolettre, je t’invite à lire <a href='https://tinyletter.com/todostudio-fr/letters/destination-to-do-studio-volume-1-numero-1-1'>le premier numéro ici</a>.
            </p>
            <p>
            <em>PS. Après ton inscription, si tu ne reçois pas rapidement un émail de confirmation, vérifie ta boîte de courrier indésirable (Junk Email).</em>
            </p>
            </body>";
        }

        private static string getTemplateEn() {
            return @"<body> 
            <p>Hello __recipientFirstName__,</p>
            <p><img src='cid:__marioPhotoCid__' height='100' width='100' align='left' hspace='10'/>
            Mario Cardinal here. I want to share my entrepreneurial journey with you. That's why I invite you to subscribe to my <a href='https://tinyletter.com/todostudio-en'>Destination To-Do Studio</a> newsletter.
            </p>
            <p>
            The newsletter is about the company I founded with __partnerName__. We spent several years searching for a simple and elegant solution to the following problem:  <em>How can we enable people to complete together a <b>common</b> to-do list?</em> And this without using the traditional approach of coordinating work with a plan and a hierarchical structure, a typical solution found with current project management software. We wanted an approach that was simpler, more open and based on the principle of self-directed teams. This to help leaders avoid the hassle of follow-ups and tough talks that occur when pushing and pulling people to take initiative.
            </p>
            <p>
            After five years of research and development (R & D), we are now ready to bring our innovation to market:  Introducing To-Do Studio, an automated assistant offered as a Software as a Service (SaaS).
            </p>
            <p>
            Here are three benefits you'll get by reading our newsletter :
                <ol>
                    <li>
                    You will gain insight into the journey of a young company that aims to get thousands of users worldwide;
                    </li>
                    <li>
                    You will follow the steps required to bring a software to market and will be able to better understand the technology;
                    </li>
                    <li>
                    Finally, you will learn how To-Do Studio is unique to empower others to take initiative.
                    </li>
                </ol>
            </p>
            <div style='text-align: center'>
                <div style='display: inline-block'>
                    <!--[if mso]>
                        <v:roundrect xmlns:v='urn:schemas-microsoft-com:vml' xmlns:w='urn:schemas-microsoft-com:office:word' 
                        href='https://tinyletter.com/todostudio-en' 
                        style='height:60px;v-text-anchor:middle;width:600px;' arcsize='15%' strokecolor='#008000' fillcolor='#008000'>
                            <w:anchorlock/>
                            <center style='color:#ffffff;font-family:sans-serif;font-size:14px;font-weight:bold;'>Yes, I want to follow your journey &rarr;</center>
                        </v:roundrect>
                    <![endif]-->
                    <a href='https://tinyletter.com/todostudio-en' 
                    style='background-color:#008000;border:1px solid #008000;border-radius:10px;color:#ffffff;display:inline-block;font-family:sans-serif;font-size:14px;font-weight:bold;line-height:60px;text-align:center;text-decoration:none;width:350px;-webkit-text-size-adjust:none;mso-hide:all;'>
                    Yes, I want to follow your journey &rarr;</a>
                </div>
            </div>
            <p>
            The newsletter is free. It is published every 2 months, 6 times a year. The subscription does not require any obligation on your part. The only information you must provide to register is your email address (it will never be shared with third parties). You can unsubscribe at any time. If you wish to preview the content of the newsletter, I invite you to read <a href='https://tinyletter.com/todostudio-en/letters/destination-to-do-studio-volume-1-issue-1'>our first issue here.</a>.
            </p>
            <p>
            <em>PS. Following registration, if you do not receive quickly a confirmation email, check your Junk Email inbox.</em>
            </p>
            </body>";
        }
    }
}
