import test
import time
import smtplib 

start_time = time.time()

fromaddr = "sairandhree.sule@niyuj.com"
# creates SMTP session 
s = smtplib.SMTP('smtp.gmail.com', 587) 

# start TLS for security 
s.starttls() 

# Authentication 
s.login(fromaddr, "Ananya12") 
    
    
    
   
test.sendEmail("sairandhree.sule+a@niyuj.com", "protected/a.txt", s)
test.sendEmail("sairandhree.sule+b@niyuj.com", "protected/b.txt", s)

 # terminating the session 
s.quit()


print("--- %s seconds ---" % (time.time() - start_time))