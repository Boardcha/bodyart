<link rel="stylesheet" href="/CSS/jquery.fancybox.min.css" />
<% if  var_sizing_type <> "finger" then %>
  In our drop-down menus for adding to cart, we show the gauge of the item you are ordering, then the length/diameter, then the ball size if it has any:
	<p>
    <select class="form-control w-auto" name="select">
      <option selected="selected">14g - 3/4 inch (5mm balls)</option>
      <option>6g - 1/2 inch (8mm balls)</option>
    </select>
	</p>
	<p>
  Jewelry is measured through a gauge system (conversion chart below). The <strong>higher</strong> the gauge number the <strong>smaller</strong> the wire is. For reference, a standard &quot;earring&quot; is usually 20 gauge. Jewelry gauges can vary between manufacturers. There is no universal regulation on companies to make a standard size in body jewelry. This is because some jewelry is made outside of the USA and measured in millimeters. And in the USA jewelry is measured in inches.</p>

  <strong>GAUGE CONVERSION CHART:</strong>
  <table class="table w-auto">
  <thead class="thead-dark">
		<tr>
		  <th class="table-cell">Gauge</th>
		  <th class="table-cell">Millimeter</th>
		  <th class="table-cell">Inch</th>
		</tr>
	</thead>
  <tr>
    <td>20g</td>
    <td>.81mm</td>
    <td></td>
  </tr>
  <tr>
    <td>18g</td>
    <td>1mm</td>
    <td></td>
  </tr>
    <tr>
      <td>16g</td>
      <td>1.2mm</td>
      <td>3/64&quot;</td>
    </tr>
    <tr>
      <td>14g</td>
      <td>1.6mm</td>
      <td>1/16&quot;</td>
    </tr>
    <tr>
      <td>12g</td>
      <td>2mm</td>
      <td>5/64&quot;</td>
    </tr>
    <tr>
      <td>10g</td>
      <td>2.4mm</td>
      <td>3/32&quot;</td>
    </tr>
    <tr>
      <td>8g</td>
      <td>3.2mm</td>
      <td>1/8&quot;</td>
    </tr>
    <tr>
      <td>6g</td>
      <td>4mm</td>
      <td>5/32&quot;</td>
    </tr>
    <tr>
      <td>4g</td>
      <td>5mm</td>
      <td>3/16&quot;</td>
    </tr>
    <tr>
      <td>2g</td>
      <td>6mm</td>
      <td>1/4&quot;</td>
    </tr>
    <tr>
      <td>0g</td>
      <td>8mm</td>
      <td>5/16&quot;</td>
    </tr>
    <tr>
      <td>00g</td>
      <td>9mm to 10mm</td>
      <td>3/8&quot;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>11mm</td>
      <td>7/16&quot;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>13mm</td>
      <td>1/2&quot;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>14mm</td>
      <td>9/16&quot;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>16mm</td>
      <td>5/8&quot;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>19mm</td>
      <td>3/4&quot;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>22mm</td>
      <td>7/8&quot;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>25mm</td>
      <td>1&quot;</td>
    </tr>
    <tr>
      <td class="text-center" colspan="3">
        <span class="btn btn-sm btn-outline-secondary" id="toggle-above-1inch">
          Show sizes above 1 inch
        </span>
      </td>
    </tr>
    <tbody style="display:none" id="above-1inch">
    <tr>
      <td>&nbsp;</td>
      <td>26.9mm</td>
      <td>1-1/16&quot;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>28.5mm</td>
      <td>1-1/8&quot;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>30.1mm</td>
      <td>1-3/16&quot;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>31.7mm</td>
      <td>1-1/4&quot;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>33.3mm</td>
      <td>1-5/16&quot;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>34.9mm</td>
      <td>1-3/8&quot;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>36.5mm</td>
      <td>1-7/16&quot;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>38.1mm</td>
      <td>1-1/2&quot;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>39.6mm</td>
      <td>1-9/16&quot;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>41.2mm</td>
      <td>1-5/8&quot;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>44.4mm</td>
      <td>1-3/4&quot;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>47.6mm</td>
      <td>1-7/8&quot;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>51mm</td>
      <td>2&quot;</td>
    </tr>
  </tbody>
  </table>

    <div class="container-fluid p-3 border border-secondary bg-light" style="border-radius: 10px;border-width:3px!important">
      <div class="row">
        <div class="col-6 col-xl-2 col-lg-3 col-md-4 col-sm-6">
          <img class="img-fluid" src="/images/measurement-guide/illustration-captive-bead-ring.png">
        </div>
        <div class="col-12 col-xl-10 col-lg-9 col-md-8 col-sm-6">
          <h4 class="m-0 p-0">Rings</h4>
          <div>Rings are an essential item that can be worn in many piercings. They come in various styles:
            <ul>
              <li>Captive bead rings - These use tension to hold a bead in place.</li>
              <li>Clickers - These use a hinge to open and close the ring.</li>
              <li>Segment rings - These have a "segment" that comes completely out of the jewelry similar to removing a bead. But it looks like a completely seamless piece of jewelry.</li>
              <li>Seamless rings - These rings require bending to open & close.</li>
            </ul>
            <h5 class="text-secondary">How to measure rings</h5>
            <ul>
              <li>Gauge - This is the <i>thickness</i> of the ring itself.</li>
              <li>Diameter/Length - Measure directly across the <i>inside</i> center of the ring.</li>
            </ul>
          </div>
        </div>
      </div>
    </div>


    <div class="container-fluid p-3 border border-secondary bg-light mt-4" style="border-radius: 10px;border-width:3px!important">
      <div class="row">
        <div class="col-6 col-xl-2 col-lg-3 col-md-4 col-sm-6">
          <img class="img-fluid" src="/images/measurement-guide/illustration-straight-barbell.png">
        </div>
        <div class="col-12 col-xl-10 col-lg-9 col-md-8 col-sm-6">
          <h4 class="m-0 p-0">Straight Barbells</h4>
          <div>Straight barbells are commonly worn in various ear projects, nipples, and tongue piercings.
            <h5 class="text-secondary mt-2">How to measure straight barbells</h5>
            <ul>
              <li>Gauge - This is the <i>thickness</i> of the post/bar itself.</li>
              <li>Length - Measure directly across the post <i>between</i> the ends.</li>
            </ul>
          </div>
        </div>
      </div>
    </div>


    <div class="container-fluid p-3 border border-secondary bg-light mt-4" style="border-radius: 10px;border-width:3px!important">
      <div class="row">
        <div class="col-6 col-xl-2 col-lg-3 col-md-4 col-sm-6">
          <img class="img-fluid" src="/images/measurement-guide/illustration-curved-barbell.png">
        </div>
        <div class="col-12 col-xl-10 col-lg-9 col-md-8 col-sm-6">
          <h4 class="m-0 p-0">Curved Barbells</h4>
          <div>Curved barbells are commonly worn in various ear projects & eyebrows.
            <h5 class="text-secondary mt-2">How to measure curved barbells</h5>
            <ul>
              <li>Gauge - This is the <i>thickness</i> of the post/bar itself.</li>
              <li>Length - Measure directly across the post <i>between</i> the ends.</li>
            </ul>
          </div>
        </div>
      </div>
    </div>


    <div class="container-fluid p-3 border border-secondary bg-light mt-4" style="border-radius: 10px;border-width:3px!important">
      <div class="row">
        <div class="col-6 col-xl-2 col-lg-3 col-md-4 col-sm-6">
          <img class="img-fluid" src="/images/measurement-guide/illustration-circular-barbell.png">
        </div>
        <div class="col-12 col-xl-10 col-lg-9 col-md-8 col-sm-6">
          <h4 class="m-0 p-0">Circular Barbells</h4>
          <div>Circular barbells are commonly worn in various ear projects, septums, nipples, and lips.
            <h5 class="text-secondary mt-2">How to measure circular barbells</h5>
            <ul>
              <li>Gauge - This is the <i>thickness</i> of the post itself.</li>
              <li>Diameter/Length - Measure directly across the <i>inside</i> center of the circular.</li>
            </ul>
          </div>
        </div>
      </div>
    </div>

  <% end if ' do not show for finger rings %>


  <% if var_sizing_type = "all" or var_sizing_type = "septum" or var_sizing_type = "captive" then %>
  <% end if ' septum %>
  
  <div class="container-fluid">
    <div class="row">
      <div class="col-6 col-xl-2 col-lg-3 col-md-4 col-sm-6 text-center">
        <a data-fancybox="size-guides" data-caption="Rings size guide" href="/images/measurement-guide/measurement-guide-rings-1000px.jpg" >
          <img class="img-fluid p-2 lazyload" data-src="/images/measurement-guide/measurement-guide-rings-1000px.jpg">
          <button class="btn btn-sm btn-secondary">Ring size comparisons</button>
        </a>
      </div>
      <div class="col-6 col-xl-2 col-lg-3 col-md-4 col-sm-6 text-center">
        <a data-fancybox="size-guides" data-caption="Ball size guide" href="/images/measurement-guide/measurement-guide-balls-1000px.jpg" >
          <img class="img-fluid p-2 lazyload" data-src="/images/measurement-guide/measurement-guide-balls-1000px.jpg">
          <button class="btn btn-sm btn-secondary">Ball size comparisons</button>
        </a>
      </div>
      <div class="col-6 col-xl-2 col-lg-3 col-md-4 col-sm-6 text-center">
        <a data-fancybox="size-guides" data-caption="Straight barbells size guide" href="/images/measurement-guide/measurement-guide-straight-barbells-1000px.jpg" >
          <img class="img-fluid p-2 lazyload" data-src="/images/measurement-guide/measurement-guide-straight-barbells-1000px.jpg">
          <button class="btn btn-sm btn-secondary">Barbell size comparisons</button>
        </a>
      </div>
      <div class="col-6 col-xl-2 col-lg-3 col-md-4 col-sm-6 text-center">
        <a data-fancybox="size-guides" data-caption="Labrets size guide" href="/images/measurement-guide/measurement-guide-labret-1000px.jpg" >
          <img class="img-fluid p-2 lazyload" data-src="/images/measurement-guide/measurement-guide-labret-1000px.jpg">
          <button class="btn btn-sm btn-secondary">Labret size comparisons</button>
        </a>
      </div>
      <div class="col-6 col-xl-2 col-lg-3 col-md-4 col-sm-6 text-center">
        <a data-fancybox="size-guides" data-caption="Circular barbells size guide" href="/images/measurement-guide/measurement-guide-circulars-1000px.jpg" >
          <img class="img-fluid p-2 lazyload " data-src="/images/measurement-guide/measurement-guide-circulars-1000px.jpg">
          <button class="btn btn-sm btn-secondary">Circular size comparisons</button>
        </a>
      </div>
      <div class="col-6 col-xl-2 col-lg-3 col-md-4 col-sm-6 text-center">
        <a data-fancybox="size-guides" data-caption="Plugs size guide" href="/images/measurement-guide/measurement-guide-plugs-1000px.jpg" >
         <img class="img-fluid p-2 lazyload" data-src="/images/measurement-guide/measurement-guide-plugs-1000px.jpg">
         <button class="btn btn-sm btn-secondary">Plug size comparisons</button>
        </a>
      </div>
      <div class="col-6 col-xl-2 col-lg-3 col-md-4 col-sm-6 text-center">
        <a data-fancybox="size-guides" data-caption="Nosescrew size guide" href="/images/measurement-guide/measurement-guide-nosescrew-1000px.jpg" >
          <img class="img-fluid p-2 lazyload" data-src="/images/measurement-guide/measurement-guide-nosescrew-1000px.jpg">
          <button class="btn btn-sm btn-secondary">Nosescrew size comparisons</button>
        </a>
      </div>
    </div>
  </div>


  <% if var_sizing_type = "all" or var_sizing_type = "finger" then %>
	<h5 class="mt-4">
		Global finger ring conversions / measurements
  </h5>
  <table class="table w-auto">
    <thead class="thead thead-dark">
      <tr>
        <th class="w-25">USA, Canada & Mexico</th>
        <th class="w-25">UK, Ireland, Australia & New Zealand </th>
        <th class="w-25">Japan</th>
        <th class="w-25">Italy, Spain & Switz</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td>4</td>
        <td>H</td>
        <td>7</td>
        <td>6.75</td>
      </tr>
      <tr>
        <td>4.5</td>
        <td>I</td>
        <td>8</td>
        <td>8</td>
      </tr>
      <tr>
        <td>5</td>
        <td>J½</td>
        <td>9</td>
        <td>9.25</td>
      </tr>
      <tr>
        <td>5.5</td>
        <td>K½</td>
        <td>10</td>
        <td>10.5</td>
      </tr>
      <tr>
        <td>6</td>
        <td>L½</td>
        <td>11</td>
        <td>11.75</td>
      </tr>
      <tr>
        <td>6.5</td>
        <td>M½</td>
        <td>13</td>
        <td>13.25</td>
      </tr>
      <tr>
        <td>7</td>
        <td>N½</td>
        <td>14</td>
        <td>14.5</td>
      </tr>
      <tr>
        <td>7.5</td>
        <td>O½</td>
        <td>15</td>
        <td>15.75</td>
      </tr>
      <tr>
        <td>8</td>
        <td>P½</td>
        <td>16</td>
        <td>17</td>
      </tr>
      <tr>
        <td>8.5</td>
        <td>Q½</td>
        <td>17</td>
        <td>18.25</td>
      </tr>
      <tr>
        <td>9</td>
        <td>R½</td>
        <td>18</td>
        <td>19.5</td>
      </tr>
      <tr>
        <td>9.5</td>
        <td>S½</td>
        <td>19</td>
        <td>20.75</td>
      </tr>
      <tr>
        <td>10</td>
        <td>T½</td>
        <td>20</td>
        <td>22</td>
      </tr>
      <tr>
        <td>10.5</td>
        <td>U½</td>
        <td>22</td>
        <td>23.25</td>
      </tr>
      <tr>
        <td>11</td>
        <td>V½</td>
        <td>23</td>
        <td>24.75</td>
      </tr>
      <tr>
        <td>11.5</td>
        <td>W½</td>
        <td>24</td>
        <td>26</td>
      </tr>
      <tr>
        <td>12</td>
        <td>X½</td>
        <td>25</td>
        <td>27.25</td>
      </tr>
      <tr>
        <td>12.5</td>
        <td>Z</td>
        <td>26</td>
        <td>28.5</td>
      </tr>
      <tr>
        <td>13</td>
        <td></td>
        <td>27</td>
        <td>29.75</td>
      </tr>
      <tr>
        <td>13.5</td>
        <td></td>
        <td></td>
        <td>31</td>
      </tr>
      <tr>
        <td>14</td>
        <td>Z3</td>
        <td></td>
        <td>32.25</td>
      </tr>
      <tr>
        <td>14.5</td>
        <td>Z4</td>
        <td></td>
        <td>33.5</td>
      </tr>
      <tr>
        <td>15</td>
        <td></td>
        <td></td>
        <td>34.75</td>
      </tr>
    </tbody>
  </table>
  <% end if %>
		<h6 class="pt-5">Enlarged example of 1&quot; on ruler</h6>
		<img class="img-fluid pb-5 lazyload dark-invert-color" data-src="images/RULER.png" alt="Ruler example" />

  