<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if var_access_level = "Social Media" or var_access_level = "Admin" then
%>
<title>Social Media | Admin</title>

<html>
    <body>
    <!--#include virtual="/admin/admin_header.asp"-->
<style>
    a {color:black}
</style>
    <div class="p-2">
    <h4 class="mb-3">Social Media Dashboard</h4>
   
          <div class="container-fluid p-0">
            <div class="row no-gutters">
                <div class="col">
                    <div class="card">
                        <div class="card-header">
                          <h5>Quick links</h5>
                        </div>
                        <div class="card-body">
                        <a href="/admin/coupons_manage.asp"><i class="fa fa-angle-right mr-2"></i>Manage coupons</a><br/>
						<a href="/admin/one-time-coupons.asp"><i class="fa fa-angle-right mr-2"></i>One time use coupons</a><br/>
						<a href="/admin/secret-sale-items.asp"><i class="fa fa-angle-right mr-2"></i>Secret sale items</a><br/>
						<a href="/admin/sliders/sliders.asp"><i class="fa fa-angle-right mr-2"></i>Manage home page sliders</a><br/>
                        <a href="https://docs.google.com/document/d/1Zet9ZgzeyqKpFHVy-hzdMh-QgC0H9xhmqRQHk7GI6t4/edit" target="_blank"><img class="mr-1" src="/images/icons/google-drive.png" height="20px">Brand guidelines</a><br/>
                        <a href="https://docs.google.com/document/d/1uH5kxrqCQI-S4uQO_FMIpHheGDkediF_-2JvSGXH9LE/edit" target="_blank"><img class="mr-1" src="/images/icons/google-drive.png" height="20px">Handbook & passwords</a><br/>
                        </div><!-- card body -->
                      </div><!-- card -->
                </div>
                <div class="col px-2">
                    
                    <div class="card">
                        <div class="card-header">
                          <h5>Reports & Analytics</h5>
                        </div>
                        <div class="card-body">                    

                            <ul class="list-unstyled">
                                   <li>
                                    <a href="/admin/power-bi/reports.asp?reportId=c204861b-779e-4e27-9d7e-4e4b6be3696a&pageName=ReportSection&reportName=Coupon Revenue">
                                        <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                        Coupon Revenue
                                        </a>
                                    </li>
                            </ul>
                        </div><!-- card body -->
                      </div><!-- card -->
                </div>
                <div class="col">
                    
                    <div class="card">
                        <div class="card-header">
                          <h5>Modeling & Visual Media</h5>
                        </div>
                        <div class="card-body">                    

                        <a href="https://drive.google.com/drive/folders/1uwENYmx79wXgHOrevhrnjG3I0FGH90gL" target="_blank"><img class="mr-1" src="/images/icons/google-drive.png" height="20px">Model photos</a><br/>
                        <a href="https://drive.google.com/drive/folders/11djLiPtOTHcOUE0kpPJ09_GRpbs75Y8n" target="_blank"><img class="mr-1" src="/images/icons/google-drive.png" height="20px">Model videos</a><br/>
                        <a href="https://drive.google.com/drive/folders/1DNLVs0WKj001N5Krr0pMjWfjYLxaMdvh" target="_blank"><img class="mr-1" src="/images/icons/google-drive.png" height="20px">Product photos</a><br/>
                        <a href="https://drive.google.com/drive/folders/10IJN3Tm_aZAjPZfDTazBhwd6Bv1gEms_" target="_blank"><img class="mr-1" src="/images/icons/google-drive.png" height="20px">Product videos</a><br/>

                        </div><!-- card body -->
                      </div><!-- card -->
                </div>
            </div><!-- row -->
        </div><!-- container -->


        </div><!-- body padding-->
    </body>
</div>


</html>
<%
else
    response.write "Access denied"
end if ' permissions
%>
