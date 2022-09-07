from flask import (
    Blueprint, flash, g, redirect, render_template, request, session, url_for,current_app
)
import random,time
def createPhoneCode(session): 
    chars=['0','1','2','3','4','5','6','7','8','9'] 
    x = random.choice(chars),random.choice(chars),random.choice(chars),random.choice(chars) 
    verifyCode = "".join(x) 
    session["phoneVerifyCode"] = {"time":int(time.time()), "code":verifyCode} 
    return verifyCode 

if __name__ == '__main__':
    createPhoneCode